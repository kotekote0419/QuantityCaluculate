using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFlowPlant3D.Services
{
    /// <summary>
    /// 敷設長（塗装面積の基礎）を「座標ベース」で算出するサービス。
    ///
    /// 方針（ユーザー合意）:
    /// - Valve / Flange / Coupling も敷設長に含める
    /// - Olet / Instrument / OrificePlate も追加
    /// - Reducer は「大径側へ100%配賦」
    /// - Tee は run / branch に分け、可能なら接続先パイプの LineTag へ高精度配賦（取れない場合は部品自身のLineTagへフォールバック）
    /// - Fastener（ガスケット等）は Port S1-S2 の距離(>0)を「厚み」として敷設長に加算（mm→m は上位で換算）
    /// </summary>
    public static class InstallLengthService
    {
        public struct Contribution
        {
            public string LineTag;
            public double LengthModelUnits; // 図面単位（後で unitToMeter を掛ける）
            public double? TargetND1;       // 寄せ先PipeのND（取れた場合のみ。MyCommands2側で数量ID解決に使う）

            public Contribution(string lineTag, double lenModel)
                : this(lineTag, lenModel, null)
            {
            }

            public Contribution(string lineTag, double lenModel, double? targetNd1)
            {
                LineTag = (lineTag ?? "");
                LengthModelUnits = lenModel;
                TargetND1 = targetNd1;
            }
        }

        /// <summary>
        /// 対象エンティティ( Pipe/Part )の敷設長寄与を返す（図面単位）。
        /// Tee/Reducerは複数寄与を返す可能性あり。
        /// 例外は握りつぶし、空配列で返す。
        /// </summary>
        public static List<Contribution> ComputeEntityContributions(DataLinksManager dlm, Transaction tr, ObjectId oid, Entity ent)
        {
            var result = new List<Contribution>();
            try
            {
                if (ent == null) return result;

                // 便利：LineTag取得（部品自身）
                string SelfLineTag()
                    => (PlantProp.GetString(dlm, oid, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号") ?? "").Trim();

                // 便利：部品自身ND（取れなければnull）
                // 便利：部品自身ND（取れなければnull）
                double? SelfNd()
                {
                    // ND1 / NominalDiameter など環境差を吸収
                    string s = (PlantProp.GetString(dlm, oid, "ND1", "NominalDiameter", "NominalDia", "呼び径", "呼び径1") ?? "").Trim();
                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d1)) return d1;

                    // Size（例: "700", "700x500x700", "1350x1000"）からも推定
                    string size = (PlantProp.GetString(dlm, oid, "Size", "NominalSize", "PartSize", "サイズ") ?? "").Trim();
                    var n = ParseNdsFromSize(size);
                    if (n.Count > 0) return n[0];

                    return null;
                }

                // Size文字列から数値列（mm/呼び径）を抽出（例: "700x500x700" -> [700,500,700]）
                List<double> ParseNdsFromSize(string size)
                {
                    var list = new List<double>();
                    if (string.IsNullOrWhiteSpace(size)) return list;

                    foreach (Match mm in Regex.Matches(size, @"\d+(\.\d+)?"))
                    {
                        if (double.TryParse(mm.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var v))
                            list.Add(v);
                    }
                    return list;
                }

                // TeeのSizeから run/branch ND を推定（接続先NDが取れない場合のフォールバック）
                (double? run, double? branch) GuessTeeNdsFromSize()
                {
                    string size = (PlantProp.GetString(dlm, oid, "Size", "NominalSize", "PartSize", "サイズ") ?? "").Trim();
                    var ns = ParseNdsFromSize(size);
                    if (ns.Count == 0) return (null, null);

                    if (ns.Count >= 3)
                    {
                        // 典型: A x B x A
                        if (Math.Abs(ns[0] - ns[2]) < 1e-9 && Math.Abs(ns[1] - ns[0]) > 1e-9)
                            return (ns[0], ns[1]);

                        // 一般: runは最大、branchは最小（保守的）
                        return (ns.Max(), ns.Min());
                    }

                    if (ns.Count == 2)
                    {
                        // 2値しか無い場合：大きい方をrun、小さい方をbranchとして扱う（Lateral等のフォールバック）
                        return (ns.Max(), ns.Min());
                    }

                    // 1値
                    return (ns[0], ns[0]);
                }

                // Pipe：直線長（Start-End）
                if (ent is Pipe pipe)
                {
                    if (TryGetPipeEnds(pipe, out var a, out var b))
                    {
                        string lt = SelfLineTag();
                        result.Add(new Contribution(lt, a.DistanceTo(b), SelfNd()));
                    }
                    return result;
                }

                // Part 系（Valve/Flange/Tee/Reducer/Connector等）
                if (!TryAsPart(ent, out var part)) return result;

                string pnpClass = (PlantProp.GetString(dlm, oid, "PnPClassName", "ClassName") ?? "").Trim();
                if (LooksLikeSupport(pnpClass, ent.GetType().Name)) return result;

                var ports = GetPortsSafe(part);
                if (ports.Count == 0) return result;

                // ----------------------------------------
                // 共通：2ポートのLineTag/ND解決（接続先Pipe優先、なければ自己）
                // ----------------------------------------
                void AddTwoPortContribution(double lenModelUnits, Port p0, Port p1, bool allowSplit)
                {
                    if (lenModelUnits <= 0) return;

                    string selfLt = SelfLineTag();

                    string lt0 = ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, p0);
                    string lt1 = ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, p1);

                    double? nd0 = ConnectionResolver.TryGetConnectedPipeNominalDiameter(dlm, tr, oid, part, p0);
                    double? nd1 = ConnectionResolver.TryGetConnectedPipeNominalDiameter(dlm, tr, oid, part, p1);

                    // 自己LineTagがあれば原則それを使う（コネクタはこのルール）
                    if (!string.IsNullOrWhiteSpace(selfLt))
                    {
                        // NDは接続先が取れればそちら、なければ自己
                        double? nd = nd0 ?? nd1 ?? SelfNd();
                        result.Add(new Contribution(selfLt, lenModelUnits, nd));
                        return;
                    }

                    // 自己LineTagが無い → 接続先へ
                    string a = (lt0 ?? "").Trim();
                    string b = (lt1 ?? "").Trim();

                    if (!string.IsNullOrWhiteSpace(a) && !string.IsNullOrWhiteSpace(b) && allowSplit && !string.Equals(a, b, StringComparison.Ordinal))
                    {
                        // LineTagが異なる場合は 1/2ずつ
                        result.Add(new Contribution(a, lenModelUnits * 0.5, nd0));
                        result.Add(new Contribution(b, lenModelUnits * 0.5, nd1));
                        return;
                    }

                    // 同一 or 片側のみ
                    string targetLt = !string.IsNullOrWhiteSpace(a) ? a : (!string.IsNullOrWhiteSpace(b) ? b : "");
                    double? targetNd = nd0 ?? nd1 ?? SelfNd();
                    result.Add(new Contribution(targetLt, lenModelUnits, targetNd));
                }

                // Tee（3ポート以上）

                if ((LooksLikeTee(pnpClass, ent.GetType().Name) || LooksLikeCross(pnpClass, ent.GetType().Name)) && ports.Count >= 3)
                {
                    // まずは堅牢判定（形状ベース：最も反対向きの2ポートをrunとする）
                    if (TryComputeTeeCrossBreakdown(dlm, tr, oid, ent, out var b))
                    {
                        AddRunContributionWithSplit(result, b.RunLenModelUnits, b.RunLineTagA, b.RunLineTagB, b.ND1);

                        if (b.BrLenModelUnits > 0) result.Add(new Contribution(b.BrLineTag, b.BrLenModelUnits, b.ND2));
                        if (b.IsCross && b.Br2LenModelUnits > 0) result.Add(new Contribution(b.Br2LineTag, b.Br2LenModelUnits, b.ND3));

                        return result;
                    }

                    // フォールバック（旧ロジック：NDベース / Ports順）
                    if (LooksLikeTee(pnpClass, ent.GetType().Name) && ports.Count >= 3)
                    {
                        var portInfo = new List<(int idx, Port port, double? nd, string lt)>();
                        for (int i = 0; i < ports.Count; i++)
                        {
                            var lt = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[i]) ?? "");
                            var nd = ConnectionResolver.TryGetConnectedPipeNominalDiameter(dlm, tr, oid, part, ports[i]);
                            portInfo.Add((i, ports[i], nd, (lt ?? "").Trim()));
                        }

                        var ndCount = portInfo.Count(x => x.nd.HasValue);
                        int run0 = 0, run1 = 1, br = 2;
                        if (ndCount >= 2)
                        {
                            var ordered = portInfo.OrderByDescending(x => x.nd ?? -1).ToList();
                            run0 = ordered[0].idx;
                            run1 = ordered[1].idx;
                            br = ordered.First(x => x.idx != run0 && x.idx != run1).idx;
                        }

                        var p0 = ports[run0].Position;
                        var p1 = ports[run1].Position;
                        var pb = ports[br].Position;

                        double runLen = p0.DistanceTo(p1);

                        Point3d j = ProjectPointToLineSegment(pb, p0, p1);
                        double branchLen = j.DistanceTo(pb);

                        string lt0 = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[run0]) ?? "");
                        string lt1 = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[run1]) ?? "");
                        double? runNd = portInfo.FirstOrDefault(x => x.idx == run0).nd ?? portInfo.FirstOrDefault(x => x.idx == run1).nd ?? SelfNd();

                        AddRunContributionWithSplit(result, runLen, lt0, lt1, runNd);

                        string brLt = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[br]) ?? "");
                        double? brNd = ConnectionResolver.TryGetConnectedPipeNominalDiameter(dlm, tr, oid, part, ports[br]) ?? SelfNd();
                        if (branchLen > 0) result.Add(new Contribution((brLt ?? "").Trim(), branchLen, brNd));

                        return result;
                    }

                    // Crossのフォールバック：Ports[0]-[1] をrun、[2],[3] をbranch扱い
                    if (ports.Count >= 4)
                    {
                        var p0 = ports[0].Position;
                        var p1 = ports[1].Position;
                        var p2 = ports[2].Position;
                        var p3 = ports[3].Position;

                        double runLen = p0.DistanceTo(p1);
                        Point3d j2 = ProjectPointToLineSegment(p2, p0, p1);
                        Point3d j3 = ProjectPointToLineSegment(p3, p0, p1);
                        double brLen = j2.DistanceTo(p2);
                        double br2Len = j3.DistanceTo(p3);

                        string lt0 = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[0]) ?? "");
                        string lt1 = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[1]) ?? "");
                        AddRunContributionWithSplit(result, runLen, lt0, lt1, SelfNd());

                        string lt2 = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[2]) ?? "");
                        string lt3 = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[3]) ?? "");
                        if (brLen > 0) result.Add(new Contribution((lt2 ?? "").Trim(), brLen, SelfNd()));
                        if (br2Len > 0) result.Add(new Contribution((lt3 ?? "").Trim(), br2Len, SelfNd()));

                        return result;
                    }
                }


                // Reducer：大径側へ100%（大径側のLineTag/NDへ）
                if (LooksLikeReducer(pnpClass, ent.GetType().Name) && ports.Count >= 2)
                {
                    var pA = ports[0].Position;
                    var pB = ports[1].Position;
                    double len = pA.DistanceTo(pB);

                    string lt0 = ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[0]);
                    string lt1 = ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, ports[1]);

                    double? nd0 = ConnectionResolver.TryGetConnectedPipeNominalDiameter(dlm, tr, oid, part, ports[0]);
                    double? nd1 = ConnectionResolver.TryGetConnectedPipeNominalDiameter(dlm, tr, oid, part, ports[1]);

                    string selfLt = SelfLineTag();
                    string targetLt = selfLt;
                    double? targetNd = nd0 ?? nd1 ?? SelfNd();

                    if (nd0.HasValue && nd1.HasValue)
                    {
                        if (nd0.Value >= nd1.Value)
                        {
                            targetLt = (lt0 ?? selfLt);
                            targetNd = nd0;
                        }
                        else
                        {
                            targetLt = (lt1 ?? selfLt);
                            targetNd = nd1;
                        }
                    }
                    else
                    {
                        targetLt = (lt0 ?? lt1 ?? selfLt);
                    }

                    result.Add(new Contribution((targetLt ?? "").Trim(), len, targetNd));
                    return result;
                }

                // Connector（ガスケット想定）：S1-S2厚みを敷設長へ
                // ※「ガスケットだけ」に限定するため、PnPClassName==P3dConnector のみ対象
                if (ent is Connector && string.Equals(pnpClass, "P3dConnector", StringComparison.OrdinalIgnoreCase))
                {
                    if (TryGetS1S2Distance(part, ports, out var thick) && thick > 0)
                    {
                        // LineTag：自己があればそれ、無ければ接続先Pipeへ（異なるなら1/2）
                        // ND：接続先PipeのNDを優先（異なる場合は各1/2に分かれる）
                        AddTwoPortContribution(thick, ports[0], ports.Count > 1 ? ports[1] : ports[0], allowSplit: true);
                        return result;
                    }
                }


                // 2ポート部品（Valve/Flange/Coupling/Instrument/OrificePlate/Olet等）：
                if (ports.Count >= 2)
                {
                    double len;

                    // Elbowは「弦長(S1-S2)」ではなく、可能なら Mid(Position) を使った経路長（S1-Mid + Mid-S2）を採用
                    if (LooksLikeElbow(pnpClass, ent.GetType().Name))
                    {
                        var mid = TryGetMidPointFromProps(dlm, oid, useCog: false);
                        if (mid.HasValue)
                            len = ports[0].Position.DistanceTo(mid.Value) + mid.Value.DistanceTo(ports[1].Position);
                        else
                            len = ports[0].Position.DistanceTo(ports[1].Position);
                    }
                    else
                    {
                        len = ports[0].Position.DistanceTo(ports[1].Position);
                    }

                    if (len > 0)
                    {
                        // 自己LineTagが無い場合は接続先Pipeへ（異なるなら1/2）
                        AddTwoPortContribution(len, ports[0], ports[1], allowSplit: true);
                    }
                }

                return result;
            }
            catch
            {
                return result;
            }
        }

        /// <summary>Fastener rowId -> （図形を走査して）S1-S2厚み（図面単位）の最大値を作る。</summary>
        public static Dictionary<int, double> BuildFastenerThicknessMapFromDrawing(Database db, DataLinksManager dlm, HashSet<int> targetRowIds)
        {
            var map = new Dictionary<int, double>();
            try
            {
                using var tr = db.TransactionManager.StartTransaction();
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId oid in btr)
                {
                    if (!oid.IsValid) continue;
                    if (!dlm.HasLinks(oid)) continue;

                    int rowId = 0;
                    try { rowId = dlm.FindAcPpRowId(oid); } catch { rowId = 0; }
                    if (rowId <= 0) continue;
                    if (!targetRowIds.Contains(rowId)) continue;

                    var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    if (!TryAsPart(ent, out var part)) continue;

                    var ports = GetPortsSafe(part);
                    if (ports.Count < 2) continue;

                    if (!TryGetS1S2Distance(part, ports, out var d)) continue;

                    if (!map.TryGetValue(rowId, out var cur) || d > cur)
                        map[rowId] = d;
                }

                tr.Commit();
            }
            catch
            {
                // ignore
            }

            return map;
        }

        public static bool TryComputeS1S2DistanceFromEntity(Entity ent, out double distanceModelUnits)
        {
            distanceModelUnits = 0;
            try
            {
                if (!TryAsPart(ent, out var part)) return false;
                var ports = GetPortsSafe(part);
                if (ports.Count < 2) return false;
                return TryGetS1S2Distance(part, ports, out distanceModelUnits);
            }
            catch { return false; }
        }

        // --------------------------
        // Internal helpers
        // --------------------------

        private static bool TryAsPart(Entity ent, out Part part)
        {
            part = null;
            try
            {
                part = ent as Part;
                return part != null;
            }
            catch { return false; }
        }
        public static bool TryGetPipeEnds(Pipe pipe, out Point3d a, out Point3d b)
        {
            a = default; b = default;
            if (pipe == null) return false;

            // 1) StartPoint / EndPoint（最優先）
            try
            {
                var piS = pipe.GetType().GetProperty("StartPoint", BindingFlags.Public | BindingFlags.Instance);
                var piE = pipe.GetType().GetProperty("EndPoint", BindingFlags.Public | BindingFlags.Instance);
                if (piS != null && piE != null)
                {
                    var vs = piS.GetValue(pipe);
                    var ve = piE.GetValue(pipe);
                    if (vs is Point3d ps && ve is Point3d pe)
                    {
                        a = ps; b = pe;
                        return true;
                    }
                }
            }
            catch { /* ignore */ }

            // 2) Ports（S1/S2 もしくは先頭2つ）
            try
            {
                // Pipe.GetPorts(PortType.Static) を反射で呼ぶ（参照差異に強くする）
                var mi = pipe.GetType().GetMethod("GetPorts", BindingFlags.Public | BindingFlags.Instance);
                if (mi != null)
                {
                    // 1引数版（PortType）を優先
                    object portsObj = null;
                    var pars = mi.GetParameters();
                    if (pars.Length == 1 && pars[0].ParameterType.Name.Contains("PortType"))
                    {
                        // PortType.Static を取得
                        var ptType = pars[0].ParameterType;
                        object staticVal = Enum.Parse(ptType, "Static");
                        portsObj = mi.Invoke(pipe, new object[] { staticVal });
                    }
                    else if (pars.Length == 0)
                    {
                        portsObj = mi.Invoke(pipe, null);
                    }

                    if (portsObj != null)
                    {
                        // IEnumerable として列挙
                        var list = new List<(string name, Point3d pos)>();
                        foreach (var p in (System.Collections.IEnumerable)portsObj)
                        {
                            if (p == null) continue;
                            string name = "";
                            try
                            {
                                var pn = p.GetType().GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
                                if (pn != null) name = (pn.GetValue(p) as string) ?? "";
                            }
                            catch { }
                            try
                            {
                                var pp = p.GetType().GetProperty("Position", BindingFlags.Public | BindingFlags.Instance);
                                var v = pp?.GetValue(p);
                                if (v is Point3d pt) list.Add((name, pt));
                            }
                            catch { }
                        }

                        if (list.Count >= 2)
                        {
                            // S1/S2 があれば優先
                            var s1 = list.FirstOrDefault(x => string.Equals(x.name, "S1", StringComparison.OrdinalIgnoreCase));
                            var s2 = list.FirstOrDefault(x => string.Equals(x.name, "S2", StringComparison.OrdinalIgnoreCase));
                            if (!s1.pos.Equals(default(Point3d)) && !s2.pos.Equals(default(Point3d)))
                            {
                                a = s1.pos; b = s2.pos;
                                return true;
                            }
                            // なければ先頭2つ
                            a = list[0].pos;
                            b = list[1].pos;
                            return true;
                        }
                    }
                }
            }
            catch { /* ignore */ }

            return false;
        }

        private static bool LooksLikeSupport(string pnpClass, string typeName)
        {
            string s = (pnpClass ?? "").ToUpperInvariant();
            string t = (typeName ?? "").ToUpperInvariant();
            return s.Contains("SUPPORT") || t.Contains("SUPPORT");
        }

        private static bool LooksLikeTee(string pnpClass, string typeName)
        {
            string s = (pnpClass ?? "").ToUpperInvariant();
            string t = (typeName ?? "").ToUpperInvariant();
            return s.Contains("TEE") || s.Contains("BRANCH") || t.Contains("TEE");
        }



        private static bool LooksLikeCross(string pnpClass, string typeName)
        {
            string s = (pnpClass ?? "").ToUpperInvariant();
            string t = (typeName ?? "").ToUpperInvariant();
            return s.Contains("CROSS") || t.Contains("CROSS");
        }

        private static bool LooksLikeElbow(string pnpClass, string typeName)
        {
            string s = ((pnpClass ?? "") + "|" + (typeName ?? "")).ToLowerInvariant();
            return s.Contains("elbow") || s.Contains("bend");
        }

        private static Point3d? TryGetMidPointFromProps(DataLinksManager dlm, ObjectId oid, bool useCog)
        {
            try
            {
                if (useCog)
                {
                    var x = PlantProp.GetDouble(dlm, oid, "COG X", "COGX");
                    var y = PlantProp.GetDouble(dlm, oid, "COG Y", "COGY");
                    var z = PlantProp.GetDouble(dlm, oid, "COG Z", "COGZ");
                    if (x.HasValue && y.HasValue && z.HasValue) return new Point3d(x.Value, y.Value, z.Value);
                }
                else
                {
                    var x = PlantProp.GetDouble(dlm, oid, "Position X", "PositionX");
                    var y = PlantProp.GetDouble(dlm, oid, "Position Y", "PositionY");
                    var z = PlantProp.GetDouble(dlm, oid, "Position Z", "PositionZ");
                    if (x.HasValue && y.HasValue && z.HasValue) return new Point3d(x.Value, y.Value, z.Value);
                }
            }
            catch { }
            return null;
        }

        private static bool LooksLikeReducer(string pnpClass, string typeName)
        {
            string s = (pnpClass ?? "").ToUpperInvariant();
            string t = (typeName ?? "").ToUpperInvariant();
            return s.Contains("REDUCER") || t.Contains("REDUCER");
        }

        private static bool LooksLikeFastener(string pnpClass, string typeName)
        {
            string s = (pnpClass ?? "").ToUpperInvariant();
            string t = (typeName ?? "").ToUpperInvariant();
            return s.Contains("GASKET") || s.Contains("BOLT") || s.Contains("FASTENER") || t.Contains("GASKET") || t.Contains("BOLT");
        }

        private static bool TryGetS1S2Distance(Part part, List<Port> ports, out double distanceModelUnits)
        {
            distanceModelUnits = 0;
            try
            {
                // 名前優先（S1/S2）
                var s1 = ports.FirstOrDefault(p => string.Equals(p.Name, "S1", StringComparison.OrdinalIgnoreCase));
                var s2 = ports.FirstOrDefault(p => string.Equals(p.Name, "S2", StringComparison.OrdinalIgnoreCase));

                if (s1 != null && s2 != null)
                {
                    distanceModelUnits = s1.Position.DistanceTo(s2.Position);
                    return distanceModelUnits > 0;
                }

                // 取れない場合は先頭2ポート
                if (ports.Count >= 2)
                {
                    distanceModelUnits = ports[0].Position.DistanceTo(ports[1].Position);
                    return distanceModelUnits > 0;
                }
            }
            catch { }
            return false;
        }

        private static List<Port> GetPortsSafe(Part part)
        {
            try
            {
                // PortType.All が使える環境を優先
                try { return ToList(part.GetPorts(PortType.All)); } catch { }
                try { return ToList(part.GetPorts(PortType.Static)); } catch { }
            }
            catch { }
            return new List<Port>();
        }

        private static List<Port> ToList(PortCollection pc)
        {
            var list = new List<Port>();
            if (pc == null) return list;
            for (int i = 0; i < pc.Count; i++) list.Add(pc[i]);
            return list;
        }

        private static Point3d ProjectPointToLineSegment(Point3d p, Point3d a, Point3d b)
        {
            // a==b
            if (a.DistanceTo(b) < 1e-9) return a;

            Vector3d ab = b - a;
            Vector3d ap = p - a;

            double t = ap.DotProduct(ab) / ab.DotProduct(ab);
            if (t < 0) t = 0;
            if (t > 1) t = 1;

            return a + ab.MultiplyBy(t);
        }


        public struct TeeCrossBreakdown
        {
            public bool IsCross;
            public double RunLenModelUnits;
            public double BrLenModelUnits;
            public double Br2LenModelUnits;

            public string RunLineTagA;
            public string RunLineTagB;
            public string BrLineTag;
            public string Br2LineTag;

            public double? ND1; // run
            public double? ND2; // br
            public double? ND3; // br2
        }

        /// <summary>
        /// Tee/Cross の run/branch を堅牢に判定し、長さ・LineTag・ND を返す。
        /// 取得不能な LineTag は null ではなく "" を返す（網羅優先）。
        /// </summary>
        public static bool TryComputeTeeCrossBreakdown(
            DataLinksManager dlm,
            Transaction tr,
            ObjectId oid,
            Entity ent,
            out TeeCrossBreakdown b)
        {
            b = default;

            try
            {
                if (ent == null) return false;
                if (ent is not Part part) return false;

                var ports = GetPortsSafe(part);
                if (ports == null || ports.Count < 3) return false;

                string pnpClass = (PlantProp.GetString(dlm, oid, "PnPClassName", "ClassName") ?? "").Trim();
                string typeName = ent.GetType().Name;
                bool isCross = LooksLikeCross(pnpClass, typeName) || ports.Count >= 4;

                int useN = isCross ? Math.Min(4, ports.Count) : 3;

                var ps = new List<(int idx, Port port, Point3d pos, string lt, double? nd)>();
                for (int i = 0; i < useN; i++)
                {
                    var p = ports[i];
                    if (p == null) continue;

                    // 接続先Pipeの LineTag / ND（取れなければ "" / null）
                    string lt = (ConnectionResolver.TryGetConnectedPipeLineTag(dlm, tr, oid, part, p) ?? "").Trim();
                    if (lt == null) lt = "";
                    double? nd = ConnectionResolver.TryGetConnectedPipeNominalDiameter(dlm, tr, oid, part, p);

                    ps.Add((i, p, p.Position, lt, nd));
                }
                if (ps.Count < 3) return false;

                // 中心点（ポート位置の平均）
                Point3d c = new Point3d(
                    ps.Average(x => x.pos.X),
                    ps.Average(x => x.pos.Y),
                    ps.Average(x => x.pos.Z)
                );

                // 中心からの単位ベクトル
                var vecs = ps.Select(x =>
                {
                    Vector3d v = x.pos - c;
                    double len = v.Length;
                    if (len < 1e-9) return (x.idx, v, 0.0);
                    return (x.idx, v / len, len);
                }).ToList();

                // runペア：最も反対向き（dot最小）な2ポート
                int runA = ps[0].idx, runB = ps[1].idx;
                double bestDot = 1.0;
                for (int i = 0; i < vecs.Count; i++)
                {
                    for (int j = i + 1; j < vecs.Count; j++)
                    {
                        double dot = vecs[i].Item2.DotProduct(vecs[j].Item2);
                        if (dot < bestDot)
                        {
                            bestDot = dot;
                            runA = vecs[i].idx;
                            runB = vecs[j].idx;
                        }
                    }
                }

                var remaining = ps.Select(x => x.idx).Where(i => i != runA && i != runB).ToList();
                if (remaining.Count < 1) return false;

                Point3d pA = ps.First(x => x.idx == runA).pos;
                Point3d pB = ps.First(x => x.idx == runB).pos;

                double runLen = pA.DistanceTo(pB);

                double BranchLen(int idx)
                {
                    Point3d pb = ps.First(x => x.idx == idx).pos;
                    Point3d j = ProjectPointToLineSegment(pb, pA, pB);
                    return j.DistanceTo(pb);
                }

                int br = remaining[0];
                int br2 = remaining.Count > 1 ? remaining[1] : -1;

                double brLen = br != -1 ? BranchLen(br) : 0.0;
                double br2Len = br2 != -1 ? BranchLen(br2) : 0.0;

                // ND 推定（接続先ND優先、なければ Size から推定）
                double? ndRunA = ps.First(x => x.idx == runA).nd;
                double? ndRunB = ps.First(x => x.idx == runB).nd;

                double? runNd = null;
                if (ndRunA.HasValue && ndRunA.Value > 0) runNd = ndRunA.Value;
                if (ndRunB.HasValue && ndRunB.Value > 0) runNd = !runNd.HasValue ? ndRunB.Value : Math.Max(runNd.Value, ndRunB.Value);

                double? brNd = br != -1 ? ps.First(x => x.idx == br).nd : null;
                double? br2Nd = br2 != -1 ? ps.First(x => x.idx == br2).nd : null;

                if (!runNd.HasValue || !brNd.HasValue || (isCross && !br2Nd.HasValue))
                {
                    string size = (PlantProp.GetString(dlm, oid, "Size", "NominalSize", "PartSize", "サイズ", "NPS") ?? "").Trim();

                    // Size文字列から数値を抽出（"700x500x700" -> [700,500,700]）
                    var nums = new List<double>();
                    foreach (Match mm in System.Text.RegularExpressions.Regex.Matches(size, @"\d+(?:\.\d+)?"))
                    {
                        if (double.TryParse(mm.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
                            if (v > 0) nums.Add(v);
                    }

                    if (nums.Count >= 2)
                    {
                        // Tee: 大きい2本=run、小さい=branch（おおまかなフォールバック）
                        // Cross: 大=run、小=branch（同じ小を2本に適用）
                        double max = nums.Max();
                        double min = nums.Min();

                        if (!runNd.HasValue) runNd = max;

                        if (!brNd.HasValue) brNd = (isCross ? min : min);

                        if (isCross && !br2Nd.HasValue) br2Nd = min;
                    }
                }

                // Cross の br/br2 順序：ND2 >= ND3 になるように並べる
                if (isCross && br2 != -1)
                {
                    double aNd = brNd ?? -1;
                    double bNd = br2Nd ?? -1;
                    if (aNd < bNd)
                    {
                        (br, br2) = (br2, br);
                        (brNd, br2Nd) = (br2Nd, brNd);
                        (brLen, br2Len) = (br2Len, brLen);
                    }
                }

                string runLtA = (ps.First(x => x.idx == runA).lt ?? "").Trim();
                string runLtB = (ps.First(x => x.idx == runB).lt ?? "").Trim();
                string brLt = br != -1 ? (ps.First(x => x.idx == br).lt ?? "").Trim() : "";
                string br2Lt = br2 != -1 ? (ps.First(x => x.idx == br2).lt ?? "").Trim() : "";

                b = new TeeCrossBreakdown
                {
                    IsCross = isCross,
                    RunLenModelUnits = Math.Max(0, runLen),
                    BrLenModelUnits = Math.Max(0, brLen),
                    Br2LenModelUnits = Math.Max(0, br2Len),

                    RunLineTagA = runLtA,
                    RunLineTagB = runLtB,
                    BrLineTag = brLt,
                    Br2LineTag = br2Lt,

                    ND1 = runNd,
                    ND2 = brNd,
                    ND3 = isCross ? br2Nd : null
                };

                return true;
            }
            catch
            {
                b = default;
                return false;
            }
        }

        private static void AddRunContributionWithSplit(
            List<Contribution> result,
            double runLenModelUnits,
            string ltA,
            string ltB,
            double? runNd)
        {
            string a = (ltA ?? "").Trim();
            string b = (ltB ?? "").Trim();

            if (runLenModelUnits <= 0) return;

            if (string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b))
            {
                result.Add(new Contribution("", runLenModelUnits, runNd));
                return;
            }

            if (!string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b))
            {
                result.Add(new Contribution(a, runLenModelUnits, runNd));
                return;
            }

            if (string.IsNullOrWhiteSpace(a) && !string.IsNullOrWhiteSpace(b))
            {
                result.Add(new Contribution(b, runLenModelUnits, runNd));
                return;
            }

            if (string.Equals(a, b, StringComparison.Ordinal))
            {
                result.Add(new Contribution(a, runLenModelUnits, runNd));
                return;
            }

            // 両端が異なる → 1/2ずつ
            result.Add(new Contribution(a, runLenModelUnits * 0.5, runNd));
            result.Add(new Contribution(b, runLenModelUnits * 0.5, runNd));
        }

        /// <summary>
        /// ConnectionManager から「接続先Pipe」の情報を取る（取れなければ null）。
        /// ※Plant3D API差分が出やすいので reflection + try/catch で安全に。
        /// </summary>
        private static class ConnectionResolver
        {
            private static object _mgr;
            private static Type _mgrType;
            private static MethodInfo _getConnsByOid;
            private static MethodInfo _getConnsByPart;

            private static bool Ensure()
            {
                try
                {
                    if (_mgr != null) return true;

                    _mgrType = typeof(ConnectionManager);

                    // singleton 取得（候補を順に試す）
                    _mgr = GetSingleton(_mgrType, "Current")
                           ?? GetSingleton(_mgrType, "Instance")
                           ?? InvokeStatic(_mgrType, "GetInstance")
                           ?? InvokeStatic(_mgrType, "GetCurrent");

                    if (_mgr == null) return false;

                    // GetConnections メソッド候補
                    _getConnsByOid = _mgrType.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                        .FirstOrDefault(m => string.Equals(m.Name, "GetConnections", StringComparison.OrdinalIgnoreCase)
                                          && m.GetParameters().Length == 1
                                          && m.GetParameters()[0].ParameterType == typeof(ObjectId));

                    _getConnsByPart = _mgrType.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                        .FirstOrDefault(m => string.Equals(m.Name, "GetConnections", StringComparison.OrdinalIgnoreCase)
                                          && m.GetParameters().Length == 1
                                          && m.GetParameters()[0].ParameterType == typeof(Part));

                    return true;
                }
                catch { return false; }
            }

            public static string TryGetConnectedPipeLineTag(DataLinksManager dlm, Transaction tr, ObjectId selfOid, Part selfPart, Port selfPort)
            {
                try
                {
                    if (!Ensure()) return "";

                    var conns = GetConnections(selfOid, selfPart);
                    if (conns == null) return "";

                    foreach (var c in EnumerateConnections(conns))
                    {
                        if (TryGetOtherPartForPort(c, selfPart, selfPort, out var otherPart, out var otherOid))
                        {
                            if (otherOid != ObjectId.Null && otherOid.IsValid)
                            {
                                // connected part's line tag
                                string lt = (PlantProp.GetString(dlm, otherOid, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号") ?? "").Trim();
                                if (!string.IsNullOrWhiteSpace(lt)) return lt;
                            }
                        }
                    }
                }
                catch { }
                return "";
            }

            public static double? TryGetConnectedPipeNominalDiameter(DataLinksManager dlm, Transaction tr, ObjectId selfOid, Part selfPart, Port selfPort)
            {
                try
                {
                    if (!Ensure()) return null;

                    var conns = GetConnections(selfOid, selfPart);
                    if (conns == null) return null;

                    foreach (var c in EnumerateConnections(conns))
                    {
                        if (TryGetOtherPartForPort(c, selfPart, selfPort, out var otherPart, out var otherOid))
                        {
                            if (otherOid != ObjectId.Null && otherOid.IsValid)
                            {
                                // ND候補（環境差があるので複数候補）
                                double? nd = PlantProp.GetDouble(dlm, otherOid,
                                    "NominalDiameter", "NominalDia", "ND", "ND1", "NOM_DIA", "NPD");
                                if (nd.HasValue && nd.Value > 0) return nd.Value;
                            }
                        }
                    }
                }
                catch { }
                return null;
            }

            private static object GetConnections(ObjectId oid, Part part)
            {
                try
                {
                    if (_getConnsByOid != null)
                        return _getConnsByOid.Invoke(_mgr, new object[] { oid });
                }
                catch { }
                try
                {
                    if (_getConnsByPart != null)
                        return _getConnsByPart.Invoke(_mgr, new object[] { part });
                }
                catch { }
                return null;
            }

            private static IEnumerable<object> EnumerateConnections(object conns)
            {
                if (conns == null) yield break;

                // ConnectionCollection / IEnumerable
                if (conns is System.Collections.IEnumerable e)
                {
                    foreach (var x in e) yield return x;
                    yield break;
                }

                // ConnectionIterator など
                var miGetEnum = conns.GetType().GetMethod("GetEnumerator", BindingFlags.Public | BindingFlags.Instance);
                if (miGetEnum != null)
                {
                    var en = miGetEnum.Invoke(conns, null) as System.Collections.IEnumerator;
                    if (en != null)
                    {
                        while (en.MoveNext()) yield return en.Current;
                    }
                }
            }

            private static bool TryGetOtherPartForPort(object conn, Part selfPart, Port selfPort, out Part otherPart, out ObjectId otherOid)
            {
                otherPart = null;
                otherOid = ObjectId.Null;
                try
                {
                    if (conn == null) return false;

                    // まず Connection から Port を2つ取る
                    var ports = GetTwoPortsFromConnection(conn);
                    if (ports.Item1 == null || ports.Item2 == null) return false;

                    var p1 = ports.Item1;
                    var p2 = ports.Item2;

                    // 自分側ポート判定：Name一致 or 位置近傍
                    bool is1Self = IsSamePort(p1, selfPort);
                    bool is2Self = IsSamePort(p2, selfPort);

                    Port otherPort = null;
                    if (is1Self && !is2Self) otherPort = p2;
                    else if (is2Self && !is1Self) otherPort = p1;
                    else
                    {
                        // どちらも不明 → 位置が近い方を自分とみなす
                        double d1 = p1.Position.DistanceTo(selfPort.Position);
                        double d2 = p2.Position.DistanceTo(selfPort.Position);
                        otherPort = (d1 <= d2) ? p2 : p1;
                    }

                    // otherPort から Part を取りたい（候補: Part / Owner / Parent）
                    otherPart = GetPortOwnerPart(otherPort);
                    if (otherPart == null)
                    {
                        // Connection から Part1/Part2 を探す
                        var parts = GetTwoPartsFromConnection(conn);
                        if (parts.Item1 != null && parts.Item2 != null)
                        {
                            // 自分と違う方
                            if (!ReferenceEquals(parts.Item1, selfPart)) otherPart = parts.Item1;
                            else otherPart = parts.Item2;
                        }
                    }

                    if (otherPart == null) return false;

                    // otherPart から ObjectId を取る（候補: ObjectId プロパティ）
                    otherOid = GetObjectIdFromPart(otherPart);
                    return true;
                }
                catch { return false; }
            }

            private static (Port, Port) GetTwoPortsFromConnection(object conn)
            {
                try
                {
                    var t = conn.GetType();
                    // 候補プロパティ
                    var p1 = GetPropAs<Port>(conn, t, "Port1") ?? GetPropAs<Port>(conn, t, "StartPort") ?? GetPropAs<Port>(conn, t, "PortA");
                    var p2 = GetPropAs<Port>(conn, t, "Port2") ?? GetPropAs<Port>(conn, t, "EndPort") ?? GetPropAs<Port>(conn, t, "PortB");
                    if (p1 != null && p2 != null) return (p1, p2);

                    // Pair プロパティ
                    var pair = t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .FirstOrDefault(pi => pi.PropertyType.Name.IndexOf("Pair", StringComparison.OrdinalIgnoreCase) >= 0);
                    if (pair != null)
                    {
                        var pv = pair.GetValue(conn);
                        if (pv != null)
                        {
                            var tp = pv.GetType();
                            p1 = GetPropAs<Port>(pv, tp, "First") ?? GetPropAs<Port>(pv, tp, "Item1");
                            p2 = GetPropAs<Port>(pv, tp, "Second") ?? GetPropAs<Port>(pv, tp, "Item2");
                            if (p1 != null && p2 != null) return (p1, p2);
                        }
                    }
                }
                catch { }
                return (null, null);
            }

            private static (Part, Part) GetTwoPartsFromConnection(object conn)
            {
                try
                {
                    var t = conn.GetType();
                    var a = GetPropAs<Part>(conn, t, "Part1") ?? GetPropAs<Part>(conn, t, "StartPart") ?? GetPropAs<Part>(conn, t, "PartA");
                    var b = GetPropAs<Part>(conn, t, "Part2") ?? GetPropAs<Part>(conn, t, "EndPart") ?? GetPropAs<Part>(conn, t, "PartB");
                    return (a, b);
                }
                catch { }
                return (null, null);
            }

            private static bool IsSamePort(Port a, Port b)
            {
                try
                {
                    if (a == null || b == null) return false;
                    if (!string.IsNullOrWhiteSpace(a.Name) && !string.IsNullOrWhiteSpace(b.Name) &&
                        string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase))
                        return true;

                    // 位置で近傍判定（mm想定で 0.01 程度）
                    return a.Position.DistanceTo(b.Position) < 0.01;
                }
                catch { return false; }
            }

            private static Part GetPortOwnerPart(Port p)
            {
                try
                {
                    if (p == null) return null;
                    var t = p.GetType();
                    return GetPropAs<Part>(p, t, "Part")
                        ?? GetPropAs<Part>(p, t, "Owner")
                        ?? GetPropAs<Part>(p, t, "Parent");
                }
                catch { return null; }
            }

            private static ObjectId GetObjectIdFromPart(Part part)
            {
                try
                {
                    if (part == null) return ObjectId.Null;
                    var t = part.GetType();
                    // 代表的な名前候補
                    var pi = t.GetProperty("ObjectId", BindingFlags.Public | BindingFlags.Instance)
                          ?? t.GetProperty("AcDbObjectId", BindingFlags.Public | BindingFlags.Instance)
                          ?? t.GetProperty("Id", BindingFlags.Public | BindingFlags.Instance);
                    if (pi != null)
                    {
                        var v = pi.GetValue(part);
                        if (v is ObjectId oid) return oid;
                    }
                }
                catch { }
                return ObjectId.Null;
            }

            private static T GetPropAs<T>(object obj, Type t, string name) where T : class
            {
                try
                {
                    var pi = t.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                    if (pi == null) return null;
                    return pi.GetValue(obj) as T;
                }
                catch { return null; }
            }

            private static object GetSingleton(Type t, string propName)
            {
                try
                {
                    var pi = t.GetProperty(propName, BindingFlags.Public | BindingFlags.Static);
                    return pi?.GetValue(null);
                }
                catch { return null; }
            }

            private static object InvokeStatic(Type t, string methodName)
            {
                try
                {
                    var mi = t.GetMethod(methodName, BindingFlags.Public | BindingFlags.Static);
                    return mi?.Invoke(null, null);
                }
                catch { return null; }
            }
        }
    }
}
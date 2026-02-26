using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFlowPlant3D.Services
{
    public static class GeometryService
    {
        public static ComponentInfo ExtractComponentInfo(DataLinksManager dlm, ObjectId oid, Entity ent)
        {
            var info = new ComponentInfo
            {
                HandleString = ent.Handle.ToString()
            };

            // EntityType は PnPClassName を優先
            var pnpClass = PlantProp.GetString(dlm, oid, "PnPClassName");
            info.EntityType = !string.IsNullOrWhiteSpace(pnpClass) ? pnpClass : ent.GetType().Name;

            // Pipe
            if (ent is Pipe)
            {
                info.Start = GetPointByProp(ent, "StartPoint");
                info.End = GetPointByProp(ent, "EndPoint");
                if (info.Start.HasValue && info.End.HasValue)
                    info.Mid = Mid(info.Start.Value, info.End.Value);

                // ND（NominalDiameter）: Pipeは PortのNDが取れない環境があるため、まずDLMの埋め込みPortプロパティ（あれば）を使う
                TryFillNominalDiameters(dlm, oid, info);

                return info;
            }

            // InlineAsset/Connector の ports（取れる範囲で）
            TryGetPorts(ent, out var ports);
            if (ports != null)
            {
                if (ports.Count >= 1) info.Start = ports[0];
                if (ports.Count >= 2) info.End = ports[1];
                if (ports.Count >= 3) info.Branch = ports[2];
                if (ports.Count >= 4) info.Branch2 = ports[3];
            }

            // Mid の決定ロジック
            // - Valve / Reducer系：COG X/Y/Z を優先
            // - その他：Position X/Y/Z を参照
            // - Midが取れない場合のみ、Start-End の中点を使用
            var pnpLower = (info.EntityType ?? "").ToLowerInvariant();
            if (pnpLower.Contains("valve") || pnpLower.Contains("reducer"))
            {
                var cx = PlantProp.GetDouble(dlm, oid, "COG X", "Cog X", "COGX");
                var cy = PlantProp.GetDouble(dlm, oid, "COG Y", "Cog Y", "COGY");
                var cz = PlantProp.GetDouble(dlm, oid, "COG Z", "Cog Z", "COGZ");
                if (cx.HasValue && cy.HasValue && cz.HasValue)
                    info.Mid = new Point3d(cx.Value, cy.Value, cz.Value);
            }
            else
            {
                var px = PlantProp.GetDouble(dlm, oid, "Position X", "Pos X", "POS X");
                var py = PlantProp.GetDouble(dlm, oid, "Position Y", "Pos Y", "POS Y");
                var pz = PlantProp.GetDouble(dlm, oid, "Position Z", "Pos Z", "POS Z");
                if (px.HasValue && py.HasValue && pz.HasValue)
                    info.Mid = new Point3d(px.Value, py.Value, pz.Value);
            }

            if (!info.Mid.HasValue && info.Start.HasValue && info.End.HasValue)
                info.Mid = Mid(info.Start.Value, info.End.Value);

            // ND（NominalDiameter）は Size からパースしない。
            // DLMの埋め込み Port プロパティ（PortName/NominalDiameter…）と、
            // DLM Relationship（PartPort: Part->Port）で辿れる Port行の両方から収集して PortName順に ND1.. へ割当てる。
            TryFillNominalDiameters(dlm, oid, info);

            return info;
        }

        private static void TryFillNominalDiameters(DataLinksManager dlm, ObjectId oid, ComponentInfo info)
        {
            int baseRowId = SafeFindRowId(dlm, oid);
            if (baseRowId <= 0) return;

            var ndByPort = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            // (A) Base rowId の GetAllProperties に埋め込まれている PortName/NominalDiameter を拾う
            try
            {
                var props = dlm.GetAllProperties(baseRowId, true);
                MergeNominalDiametersFromProperties(props, ndByPort);
            }
            catch { }

            // (B) Relationship: PartPort (Part -> Port) で Port行に辿れるなら、Port行からも拾ってマージ
            try
            {
                IEnumerable related = null;
                try
                {
                    // 最も一般的なオーバーロード
                    related = dlm.GetRelatedRowIds("PartPort", "Part", baseRowId, "Port");
                }
                catch
                {
                    // 別シグネチャ環境向け（念のため reflection）
                    related = InvokeGetRelatedRowIdsFallback(dlm, baseRowId);
                }

                if (related != null)
                {
                    foreach (var ridObj in related)
                    {
                        int rid;
                        try { rid = Convert.ToInt32(ridObj); }
                        catch { continue; }
                        if (rid <= 0) continue;

                        try
                        {
                            var p2 = dlm.GetAllProperties(rid, true);
                            MergeNominalDiametersFromProperties(p2, ndByPort);
                        }
                        catch { }
                    }
                }
            }
            catch { }

            if (ndByPort.Count == 0) return;

            // PortName順（S1,S2,S3...）で ND1.. に格納
            var ordered = new List<(string Port, double Nd)>();
            foreach (var kv in ndByPort)
                ordered.Add((kv.Key, kv.Value));

            ordered.Sort((a, b) => ComparePortName(a.Port, b.Port));

            if (ordered.Count >= 1) info.ND1 = ordered[0].Nd;
            if (ordered.Count >= 2) info.ND2 = ordered[1].Nd;
            if (ordered.Count >= 3) info.ND3 = ordered[2].Nd;
        }

        private static void MergeNominalDiametersFromProperties(List<KeyValuePair<string, string>> props, Dictionary<string, double> ndByPort)
        {
            if (props == null || props.Count == 0) return;

            string currentPort = null;
            double? currentNd = null;

            // 典型的には PortName → NominalDiameter → NominalUnit ... の順で現れる（繰返し）
            for (int i = 0; i < props.Count; i++)
            {
                var k = props[i].Key ?? "";
                var v = props[i].Value ?? "";

                if (k.Equals("PortName", StringComparison.OrdinalIgnoreCase))
                {
                    // 直前のまとまりを確定
                    CommitCurrent();
                    currentPort = v?.Trim();
                    currentNd = null;
                    continue;
                }

                if (k.Equals("NominalDiameter", StringComparison.OrdinalIgnoreCase))
                {
                    if (TryParseDouble(v, out var nd))
                        currentNd = nd;
                    continue;
                }
            }

            // 最後のまとまり
            CommitCurrent();

            void CommitCurrent()
            {
                if (string.IsNullOrWhiteSpace(currentPort)) return;
                if (!currentNd.HasValue) return;

                // すでに同Portがある場合は、値が入っていない/0 のものを優先的に置換
                if (!ndByPort.ContainsKey(currentPort) || ndByPort[currentPort] <= 0.0)
                    ndByPort[currentPort] = currentNd.Value;
            }
        }

        private static bool TryParseDouble(string s, out double v)
        {
            v = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            // Plantプロパティは整数文字列が多いが、小数も許容
            return double.TryParse(s.Trim(),
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture, out v)
                || double.TryParse(s.Trim(),
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.CurrentCulture, out v);
        }

        private static int ComparePortName(string a, string b)
        {
            // S1,S2,S3... を数値で比較。その他は文字列比較。
            int na = ExtractPortNumber(a);
            int nb = ExtractPortNumber(b);
            if (na >= 0 && nb >= 0) return na.CompareTo(nb);
            if (na >= 0) return -1;
            if (nb >= 0) return 1;
            return string.Compare(a ?? "", b ?? "", StringComparison.OrdinalIgnoreCase);
        }

        private static int ExtractPortNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return -1;
            s = s.Trim();
            if (s.Length < 2) return -1;
            if (!(s[0] == 'S' || s[0] == 's')) return -1;
            if (int.TryParse(s.Substring(1), out var n)) return n;
            return -1;
        }

        private static IEnumerable InvokeGetRelatedRowIdsFallback(DataLinksManager dlm, int baseRowId)
        {
            try
            {
                var t = dlm.GetType();
                var ms = t.GetMethods(BindingFlags.Public | BindingFlags.Instance);
                foreach (var m in ms)
                {
                    if (!string.Equals(m.Name, "GetRelatedRowIds", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var ps = m.GetParameters();
                    if (ps.Length != 4) continue;
                    if (ps[0].ParameterType != typeof(string)) continue;
                    if (ps[1].ParameterType != typeof(string)) continue;
                    if (ps[2].ParameterType != typeof(int)) continue;
                    if (ps[3].ParameterType != typeof(string)) continue;

                    var r = m.Invoke(dlm, new object[] { "PartPort", "Part", baseRowId, "Port" });
                    return r as IEnumerable;
                }
            }
            catch { }
            return null;
        }

        private static int SafeFindRowId(DataLinksManager dlm, ObjectId oid)
        {
            try { return dlm.FindAcPpRowId(oid); }
            catch { return -1; }
        }

        private static bool TryGetPorts(Entity ent, out List<Point3d> ports)
        {
            ports = new List<Point3d>();
            try
            {
                if (ent is Pipe pipe)
                {
                    // Pipe は GetPorts が無い場合があるため、Start/End を Port(S1/S2) として扱う
                    var s = GetPointByProp(pipe, "StartPoint");
                    var e = GetPointByProp(pipe, "EndPoint");
                    if (s.HasValue) ports.Add(s.Value);
                    if (e.HasValue) ports.Add(e.Value);
                    return ports.Count > 0;
                }

                if (ent is PipeInlineAsset asset)
                {
                    var pc = asset.GetPorts(PortType.Static);
                    for (int i = 0; i < pc.Count; i++) ports.Add(pc[i].Position);
                    return ports.Count > 0;
                }

                if (ent is Connector connector)
                {
                    var pc = connector.GetPorts(PortType.Static);
                    for (int i = 0; i < pc.Count; i++) ports.Add(pc[i].Position);
                    return ports.Count > 0;
                }
            }
            catch { }
            return false;
        }

        private static Point3d Mid(Point3d a, Point3d b)
            => new Point3d((a.X + b.X) / 2.0, (a.Y + b.Y) / 2.0, (a.Z + b.Z) / 2.0);

        private static Point3d? GetPointByProp(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) return null;
                var v = pi.GetValue(obj);
                if (v is Point3d p) return p;
            }
            catch { }
            return null;
        }

        // ===== ND list helpers (dynamic ND columns) =====

        /// <summary>
        /// ObjectId(実体)から、PortName順(S1,S2,S3...)の呼び径リストを返す。
        /// </summary>

        /// <summary>
        /// Entity から取得できる範囲で Port 座標（S1,S2,S3...順）を返す。
        /// ※座標のみ。PortName は現状 S1,S2... とみなして CSV 列名に反映する。
        /// </summary>
        public static List<Point3d> GetPorts(Entity ent)
        {
            TryGetPorts(ent, out var ports);
            return ports ?? new List<Point3d>();
        }

        public static List<string> GetNominalDiametersList(DataLinksManager dlm, ObjectId oid, Editor ed = null)
        {
            try
            {
                int baseRowId = SafeFindRowId(dlm, oid);
                return GetNominalDiametersListByRowId(dlm, baseRowId, ed);
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\n[UFLOW][DBG] GetNominalDiametersList(ObjectId) failed: {ex.GetType().Name}:{ex.Message}");
                return new List<string>();
            }
        }

        /// <summary>
        /// RowId(=PnPID)から、PortName順の呼び径リストを返す。
        /// (Pipe/Part だけでなく、FastenerRow(PnPDatabase行)にも使える)
        /// </summary>
        public static List<string> GetNominalDiametersListByRowId(DataLinksManager dlm, int baseRowId, Editor ed = null)
        {
            try
            {
                var list = new List<string>();
                if (dlm == null || baseRowId <= 0) return list;

                var ndByPort = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

                // (A) Base row properties
                try
                {
                    var props = dlm.GetAllProperties(baseRowId, true);
                    MergeNominalDiametersFromProperties(props, ndByPort);
                }
                catch { }

                // (B) Relationship PartPort: Part -> Port
                try
                {
                    IEnumerable related = null;
                    try
                    {
                        related = dlm.GetRelatedRowIds("PartPort", "Part", baseRowId, "Port");
                    }
                    catch
                    {
                        related = InvokeGetRelatedRowIdsFallback(dlm, baseRowId);
                    }

                    if (related != null)
                    {
                        foreach (var ridObj in related)
                        {
                            int rid;
                            try { rid = Convert.ToInt32(ridObj); }
                            catch { continue; }
                            if (rid <= 0) continue;

                            try
                            {
                                var p2 = dlm.GetAllProperties(rid, true);
                                MergeNominalDiametersFromProperties(p2, ndByPort);
                            }
                            catch { }
                        }
                    }
                }
                catch { }

                if (ndByPort.Count == 0) return list;

                var ordered = new List<(string Port, double Nd)>();
                foreach (var kv in ndByPort) ordered.Add((kv.Key, kv.Value));
                ordered.Sort((a, b) => ComparePortName(a.Port, b.Port));

                foreach (var it in ordered)
                {
                    var nd = it.Nd;
                    var s = ((nd % 1.0) == 0.0)
                        ? ((int)nd).ToString(CultureInfo.InvariantCulture)
                        : nd.ToString(CultureInfo.InvariantCulture);
                    list.Add(s);
                }

                return list;
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage($"\n[UFLOW][DBG] GetNominalDiametersListByRowId failed: rowId={baseRowId} ex={ex.GetType().Name}:{ex.Message}");
                return new List<string>();
            }
        }

        /// <summary>
        /// FastenerRow(PnPDatabase行)向け：Port ND が見つからない場合 Size をフォールバックとして使う。
        /// 取得できた ND が1つしかない場合でも、Gasketは2端として補完する。
        /// </summary>
        public static List<string> GetNominalDiametersListForFastenerRow(
            DataLinksManager dlm,
            int fastenerRowId,
            string sizeFallback,
            string fastenerClassName,
            Editor ed = null)
        {
            var list = GetNominalDiametersListByRowId(dlm, fastenerRowId, ed);

            int desired = 2;
            if (!string.IsNullOrWhiteSpace(fastenerClassName))
            {
                if (fastenerClassName.Equals("BoltSet", StringComparison.OrdinalIgnoreCase)) desired = 1;
                else if (fastenerClassName.Equals("Gasket", StringComparison.OrdinalIgnoreCase)) desired = 2;
            }

            // If we already found ports but only one for gasket, try to complement
            if (desired == 2 && list.Count == 1)
            {
                // If Size is numeric and equals the found ND, duplicate it.
                var s = (sizeFallback ?? "").Trim();
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var ndVal))
                {
                    var ndStr = ((ndVal % 1.0) == 0.0) ? ((int)ndVal).ToString(CultureInfo.InvariantCulture) : ndVal.ToString(CultureInfo.InvariantCulture);
                    if (list[0] == ndStr) list.Add(ndStr);
                }
            }

            if (list.Count > 0) return list;

            // Fallback by Size
            string ndStr2 = "";
            if (!string.IsNullOrWhiteSpace(sizeFallback))
            {
                var s = sizeFallback.Trim();
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var ndVal))
                {
                    ndStr2 = ((ndVal % 1.0) == 0.0) ? ((int)ndVal).ToString(CultureInfo.InvariantCulture) : ndVal.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    var m = System.Text.RegularExpressions.Regex.Match(s, @"(\d+(?:\.\d+)?)");
                    if (m.Success && double.TryParse(m.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out ndVal))
                    {
                        ndStr2 = ((ndVal % 1.0) == 0.0) ? ((int)ndVal).ToString(CultureInfo.InvariantCulture) : ndVal.ToString(CultureInfo.InvariantCulture);
                    }
                }
            }

            if (string.IsNullOrEmpty(ndStr2))
            {
                ed?.WriteMessage($"\n[UFLOW][DBG] FastenerRow ND missing: rowId={fastenerRowId} Size='{sizeFallback}' Class='{fastenerClassName}'");
                return new List<string>();
            }

            var outList = new List<string>();
            for (int i = 0; i < desired; i++) outList.Add(ndStr2);
            return outList;
        }

    }
}
// QuantityKeyBuilder.cs
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFlowPlant3D.Services
{
    /// <summary>
    /// 数量IDキー生成：
    /// - Pipe: MaterialCode + 施工方法 + Size
    /// - Elbow: PartFamilyLongDesc + 施工方法 + Size + Angle
    /// - Tee:   PartFamilyLongDesc + 施工方法 + RunND x BranchND（Port index: 0/1=run, 2=branch）
    /// - Joint系: PartFamilyLongDesc + 施工方法 + Size
    /// - その他: ItemCode + 施工方法 （無ければ Desc + 施工方法 + Size）
    /// </summary>
    public static class QuantityKeyBuilder
    {
        public static string BuildKey(DataLinksManager dlm, ObjectId oid, Entity ent)
        {
            string install = PlantProp.GetString(dlm, oid, "施工方法", "Installation", "INSTALLATION");
            string sizeRaw = PlantProp.GetString(dlm, oid, "サイズ", "Size", "NPS");
            string size = NormalizeSize(sizeRaw);

            // -------------------------
            // Pipe
            // -------------------------
            if (ent is Pipe)
            {
                string mat = PlantProp.GetString(dlm, oid, "材料コード", "MaterialCode", "MAT_CODE");
                return $"PIPE|{mat}|{install}|{size}";
            }

            // -------------------------
            // Part（InlineAsset/ConnectorなどはPart継承が多い）
            // -------------------------
            if (ent is Part part)
            {
                string typeName = ent.GetType().Name;
                string desc = PlantProp.GetString(dlm, oid, "部品仕様詳細", "PartFamilyLongDesc", "LONG_DESC");
                string item = PlantProp.GetString(dlm, oid, "項目コード", "ItemCode", "ITEM_CODE");

                // ---- Elbow：角度をキーへ
                if (typeName.IndexOf("Elbow", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string ang = NormalizeAngle(PlantProp.GetString(dlm, oid, "角度", "Angle", "PathAngle"));
                    return $"ELBOW|{desc}|{install}|{size}|{ang}";
                }

                // ---- Tee：枝管径までキーへ（Port indexで判定）
                if (typeName.IndexOf("Tee", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string runNd = "";
                    string brNd = "";

                    // 1) Portから取れるならそれを優先
                    try
                    {
                        var ports = PortUtil.ToPortArray(part.GetPorts(PortType.Static));
                        if (ports.Length >= 3)
                        {
                            // ports[0], ports[1] = run、ports[2] = branch（あなたの仕様）
                            runNd = NormalizeSize(ToInvariantString(ports[0].NominalDiameter));
                            brNd = NormalizeSize(ToInvariantString(ports[2].NominalDiameter));
                        }
                    }
                    catch
                    {
                        // NominalDiameter が無い/例外等は無視してfallbackへ
                    }

                    // 2) fallback: "100x50" 形式をSize文字列から分解
                    if (string.IsNullOrEmpty(runNd) || string.IsNullOrEmpty(brNd))
                    {
                        var x = (sizeRaw ?? "").Replace(" ", "").Split('x', 'X');
                        if (x.Length >= 2)
                        {
                            runNd = NormalizeSize(x[0]);
                            brNd = NormalizeSize(x[1]);
                        }
                        else
                        {
                            runNd = size;
                            brNd = "";
                        }
                    }

                    return $"TEE|{desc}|{install}|{runNd}x{brNd}";
                }

                // ---- Joint系（必要に応じて増やしてOK）
                if (typeName.IndexOf("Flange", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    typeName.IndexOf("Reducer", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    typeName.IndexOf("Coupling", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    typeName.IndexOf("BlindFlange", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return $"JOINT|{desc}|{install}|{size}";
                }

                // ---- その他Asset（Fastener含む）：基本は ItemCode + install
                if (!string.IsNullOrEmpty(item))
                    return $"ASSET|{item}|{install}";

                // ItemCode が無ければ desc + install + size
                if (!string.IsNullOrEmpty(desc))
                    return $"ASSET|{desc}|{install}|{size}";

                return $"ASSET|{typeName}|{install}|{size}";
            }

            // -------------------------
            // 最終fallback
            // -------------------------
            return $"UNKNOWN|{install}|{oid.Handle}";
        }

        private static string NormalizeSize(string s)
        {
            s = (s ?? "").Trim();
            if (string.IsNullOrEmpty(s)) return "";

            // "10.0" -> "10"
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
            {
                if (Math.Abs(d - Math.Round(d)) < 1e-9)
                    return ((int)Math.Round(d)).ToString(CultureInfo.InvariantCulture);

                return d.ToString("0.###", CultureInfo.InvariantCulture);
            }

            // "100 x 50" -> "100x50"
            return s.Replace(" ", "");
        }

        private static string NormalizeAngle(string s)
        {
            s = (s ?? "").Trim();
            if (string.IsNullOrEmpty(s)) return "";

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
            {
                if (Math.Abs(d - Math.Round(d)) < 1e-9)
                    return ((int)Math.Round(d)).ToString(CultureInfo.InvariantCulture);

                return d.ToString("0.###", CultureInfo.InvariantCulture);
            }

            return s.Replace(" ", "");
        }

        /// <summary>
        /// Plant3DのGetPorts戻り値はバージョン/参照で Port[] とは限らないため、
        /// IEnumerable を走査して Port[] に落とす。
        /// </summary>
        private static class PortUtil
        {
            public static Port[] ToPortArray(object portsObj)
            {
                if (portsObj == null) return Array.Empty<Port>();

                // PortCollection 等は IEnumerable を実装していることが多い
                if (portsObj is IEnumerable e)
                {
                    var list = new List<Port>();
                    foreach (var x in e)
                    {
                        if (x is Port p) list.Add(p);
                    }
                    return list.ToArray();
                }

                return Array.Empty<Port>();
            }
        }

        /// <summary>
        /// IFormattable を見てから文字列化
        /// </summary>
        private static string ToInvariantString(object v)
        {
            if (v == null) return "";
            if (v is IFormattable f) return f.ToString(null, CultureInfo.InvariantCulture);
            return v.ToString();
        }

    }
}

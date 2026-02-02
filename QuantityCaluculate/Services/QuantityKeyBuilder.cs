using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFlowPlant3D.Services
{
    /// <summary>
    /// 数量IDキー生成：
    /// - Pipe:     "PIPE|MaterialCode|InstallType|Size"
    /// - Elbow:    "ELBOW|PartFamilyLongDesc(or ItemCode)|InstallType|Size|Angle"
    /// - Tee:      "TEE|PartFamilyLongDesc(or ItemCode)|InstallType|RunNDxBranchND"
    /// - Fastener: "FASTENER|ItemCode(or Desc)|InstallType|Size"
    /// - その他:    "ASSET|ItemCode(or Desc)|InstallType|Size"
    /// </summary>
    public static class QuantityKeyBuilder
    {
        public static string BuildKey(DataLinksManager dlm, ObjectId oid, Entity ent)
        {
            string install = PlantProp.GetString(dlm, oid, "施工方法", "Installation", "INSTALLATION");
            string sizeRaw = PlantProp.GetString(dlm, oid, "サイズ", "Size", "NPS");
            string size = NormalizeSize(sizeRaw);

            // Pipe
            if (ent is Pipe)
            {
                string mat = PlantProp.GetString(dlm, oid, "材料コード", "MaterialCode", "MAT_CODE");
                return $"PIPE|{mat}|{install}|{size}";
            }

            string typeName = ent.GetType().Name;
            string desc = PlantProp.GetString(dlm, oid, "部品仕様詳細", "PartFamilyLongDesc", "LONG_DESC", "ShortDescription");
            string item = PlantProp.GetString(dlm, oid, "項目コード", "ItemCode", "ITEM_CODE");
            string keyBase = !string.IsNullOrWhiteSpace(desc) ? desc : item;
            if (string.IsNullOrWhiteSpace(keyBase)) keyBase = typeName;

            // Fastener
            if (IsFastener(dlm, oid, typeName))
            {
                string fastBase = !string.IsNullOrWhiteSpace(item) ? item : desc;
                if (string.IsNullOrWhiteSpace(fastBase)) fastBase = typeName;
                return $"FASTENER|{fastBase}|{install}|{size}";
            }

            // Part（InlineAsset/ConnectorなどはPart継承が多い）
            if (ent is Part part)
            {
                // Elbow
                if (typeName.IndexOf("Elbow", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string ang = NormalizeAngle(PlantProp.GetString(dlm, oid, "角度", "Angle", "PathAngle"));
                    return $"ELBOW|{keyBase}|{install}|{size}|{ang}";
                }

                // Tee
                if (typeName.IndexOf("Tee", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var (runNd, brNd) = ExtractRunBranchNd(part, sizeRaw, size);
                    return $"TEE|{keyBase}|{install}|{runNd}x{brNd}";
                }

                // その他Asset
                return $"ASSET|{keyBase}|{install}|{size}";
            }

            // 最終fallback
            return $"ASSET|{keyBase}|{install}|{size}";
        }

        private static (string runNd, string brNd) ExtractRunBranchNd(object part, string sizeRaw, string sizeFallback)
        {
            string runNd = "";
            string brNd = "";

            // 1) Portから取れるならそれを優先
            try
            {
                var ports = PortUtil.ToPortObjects(TryInvokeGetPorts(part, PortType.Static));
                if (ports.Count >= 3)
                {
                    runNd = NormalizeSize(ToInvariantString(TryGetPropertyValue(ports[0], "NominalDiameter")));
                    brNd = NormalizeSize(ToInvariantString(TryGetPropertyValue(ports[2], "NominalDiameter")));
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
                    runNd = sizeFallback;
                    brNd = "";
                }
            }

            return (runNd, brNd);
        }

        private static bool IsFastener(DataLinksManager dlm, ObjectId oid, string typeName)
        {
            if (typeName.IndexOf("Fastener", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (typeName.IndexOf("Gasket", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (typeName.IndexOf("Bolt", StringComparison.OrdinalIgnoreCase) >= 0) return true;

            var partType = PlantProp.GetString(dlm, oid, "PartType", "PartCategory", "ComponentType");
            if (partType.IndexOf("Fastener", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (partType.IndexOf("Gasket", StringComparison.OrdinalIgnoreCase) >= 0) return true;

            return false;
        }

        private static string NormalizeSize(string s)
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

        private static string ToInvariantString(object v)
        {
            if (v == null) return "";
            if (v is IFormattable f) return f.ToString(null, CultureInfo.InvariantCulture);
            return v.ToString();
        }

        private static object TryInvokeGetPorts(object ent, PortType portType)
        {
            try
            {
                var mi = ent.GetType().GetMethod("GetPorts", BindingFlags.Public | BindingFlags.Instance, null, new[] { typeof(PortType) }, null);
                if (mi == null) return null;
                return mi.Invoke(ent, new object[] { portType });
            }
            catch
            {
                return null;
            }
        }

        private static object TryGetPropertyValue(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                return pi?.GetValue(obj);
            }
            catch
            {
                return null;
            }
        }

        private static class PortUtil
        {
            public static List<object> ToPortObjects(object portsObj)
            {
                var list = new List<object>();
                if (portsObj == null) return list;

                if (portsObj is IEnumerable e)
                {
                    foreach (var x in e)
                    {
                        if (x != null) list.Add(x);
                    }
                }

                return list;
            }
        }
    }
}

using System;
using System.Collections.Specialized;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;

namespace UFlowPlant3D.Services
{
    public static class QuantityKeyProp
    {
        // DWG側に「無い場合に作る」保存先
        private const string XDICT_KEY = "UFLOW";
        private const string XREC_KEY = "QTYKEY";
        private const string XREC_LEGACY_KEY = "UFLOW:数量ID";

        /// <summary>
        /// 数量キー（文字列）を取得：
        /// 1) Plant3Dプロパティ（数量ID/QuantityID/QTY_ID）を優先
        /// 2) 無ければDWGのXRecordから
        /// </summary>
        public static string Get(DataLinksManager dlm, Transaction tr, ObjectId oid)
        {
            var s = PlantProp.GetString(dlm, oid, "数量ID", "QuantityID", "QTY_ID");
            if (!string.IsNullOrWhiteSpace(s)) return s;

            var fromDict = XRecordUtil.ReadString(tr, oid, XDICT_KEY, XREC_KEY);
            if (!string.IsNullOrWhiteSpace(fromDict)) return fromDict;

            var legacy = XRecordUtil.ReadLegacyString(tr, oid, XREC_LEGACY_KEY);
            if (!string.IsNullOrWhiteSpace(legacy))
            {
                XRecordUtil.WriteString(tr, oid, XDICT_KEY, XREC_KEY, legacy);
                return legacy;
            }

            return "";
        }

        /// <summary>
        /// 数量キー（文字列）を設定：
        /// 1) Plant3Dプロパティに書ければ書く
        /// 2) 書けなければDWGのXRecordへ保存（=「無い場合は作る」）
        /// </summary>
        public static void Set(DataLinksManager dlm, Transaction tr, ObjectId oid, string key)
        {
            key ??= "";

            if (TrySetPlantProp(dlm, oid, "数量ID", key) ||
                TrySetPlantProp(dlm, oid, "QuantityID", key) ||
                TrySetPlantProp(dlm, oid, "QTY_ID", key))
            {
                return;
            }

            XRecordUtil.WriteString(tr, oid, XDICT_KEY, XREC_KEY, key);
        }

        private static bool TrySetPlantProp(DataLinksManager dlm, ObjectId oid, string propName, string value)
        {
            try
            {
                int? rowId = PlantProp.TryGetRowIdPublic(dlm, oid);
                var ppoid = PlantProp.TryGetPpObjectId(dlm, oid);

                if (rowId.HasValue && InvokeSet(dlm, rowId.Value, propName, value)) return true;
                if (InvokeSet(dlm, oid, propName, value)) return true;
                if (ppoid != null && InvokeSet(dlm, ppoid, propName, value)) return true;
            }
            catch { }

            return false;
        }

        private static bool InvokeSet(DataLinksManager dlm, object firstArg, string propName, string value)
        {
            try
            {
                foreach (var mi in dlm.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance))
                {
                    if (mi.Name != "SetProperties") continue;
                    var ps = mi.GetParameters();
                    if (ps.Length != 3) continue;

                    if (!ps[0].ParameterType.IsInstanceOfType(firstArg)) continue;

                    var namesObj = BuildNames(ps[1].ParameterType, propName);
                    var valuesObj = BuildValues(ps[2].ParameterType, value);

                    if (namesObj == null || valuesObj == null) continue;

                    mi.Invoke(dlm, new object[] { firstArg, namesObj, valuesObj });
                    return true;
                }
            }
            catch { }

            return false;
        }

        private static object BuildNames(Type targetType, string propName)
        {
            if (targetType == typeof(StringCollection))
            {
                return new StringCollection { propName };
            }

            if (targetType == typeof(string[]))
            {
                return new[] { propName };
            }

            if (targetType == typeof(object[]))
            {
                return new object[] { propName };
            }

            return null;
        }

        private static object BuildValues(Type targetType, string value)
        {
            if (targetType == typeof(StringCollection))
            {
                return new StringCollection { value };
            }

            if (targetType == typeof(string[]))
            {
                return new[] { value };
            }

            if (targetType == typeof(object[]))
            {
                return new object[] { value };
            }

            return null;
        }
    }

    internal static class XRecordUtil
    {
        public static string SafeDictKey(string key)
        {
            if (string.IsNullOrEmpty(key)) return "_";

            var chars = key.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                var c = chars[i];
                bool ok = (c >= 'a' && c <= 'z') ||
                          (c >= 'A' && c <= 'Z') ||
                          (c >= '0' && c <= '9') ||
                          c == '_';
                if (!ok) chars[i] = '_';
            }

            return new string(chars);
        }

        public static string ReadString(Transaction tr, ObjectId entId, string dictName, string recordName)
        {
            var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
            if (ent == null || ent.ExtensionDictionary.IsNull) return null;

            var extDict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (extDict == null) return null;

            var safeDictName = SafeDictKey(dictName);
            if (!extDict.Contains(safeDictName)) return null;

            var uflowDict = tr.GetObject(extDict.GetAt(safeDictName), OpenMode.ForRead) as DBDictionary;
            if (uflowDict == null) return null;

            var safeRecordName = SafeDictKey(recordName);
            if (!uflowDict.Contains(safeRecordName)) return null;

            var xr = tr.GetObject(uflowDict.GetAt(safeRecordName), OpenMode.ForRead) as Xrecord;
            if (xr?.Data == null) return null;

            var arr = xr.Data.AsArray();
            if (arr == null || arr.Length == 0) return null;

            return arr[0].Value as string;
        }

        public static void WriteString(Transaction tr, ObjectId entId, string dictName, string recordName, string value)
        {
            try
            {
                var ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                if (ent == null) return;

                if (ent.ExtensionDictionary.IsNull)
                    ent.CreateExtensionDictionary();

                var extDict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
                if (extDict == null) return;

                var safeDictName = SafeDictKey(dictName);
                DBDictionary uflowDict;
                if (extDict.Contains(safeDictName))
                {
                    uflowDict = tr.GetObject(extDict.GetAt(safeDictName), OpenMode.ForWrite) as DBDictionary;
                }
                else
                {
                    uflowDict = new DBDictionary();
                    extDict.SetAt(safeDictName, uflowDict);
                    tr.AddNewlyCreatedDBObject(uflowDict, true);
                }

                if (uflowDict == null) return;

                var safeRecordName = SafeDictKey(recordName);
                Xrecord xr;
                if (uflowDict.Contains(safeRecordName))
                {
                    xr = tr.GetObject(uflowDict.GetAt(safeRecordName), OpenMode.ForWrite) as Xrecord;
                }
                else
                {
                    xr = new Xrecord();
                    uflowDict.SetAt(safeRecordName, xr);
                    tr.AddNewlyCreatedDBObject(xr, true);
                }

                xr.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, value ?? ""));
            }
            catch
            {
                // Avoid command failure even if dictionary operations fail.
            }
        }

        public static string ReadLegacyString(Transaction tr, ObjectId entId, string key)
        {
            var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
            if (ent == null || ent.ExtensionDictionary.IsNull) return null;

            var dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (dict == null) return null;

            var safeKey = SafeDictKey(key);
            if (!dict.Contains(safeKey)) return null;

            var xr = tr.GetObject(dict.GetAt(safeKey), OpenMode.ForRead) as Xrecord;
            if (xr?.Data == null) return null;

            var arr = xr.Data.AsArray();
            if (arr == null || arr.Length == 0) return null;

            return arr[0].Value as string;
        }

        public static void WriteLegacyString(Transaction tr, ObjectId entId, string key, string value)
        {
            try
            {
                var ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
                if (ent == null) return;

                if (ent.ExtensionDictionary.IsNull)
                    ent.CreateExtensionDictionary();

                var dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
                if (dict == null) return;

                var safeKey = SafeDictKey(key);
                Xrecord xr;
                if (dict.Contains(safeKey))
                {
                    xr = tr.GetObject(dict.GetAt(safeKey), OpenMode.ForWrite) as Xrecord;
                }
                else
                {
                    xr = new Xrecord();
                    dict.SetAt(safeKey, xr);
                    tr.AddNewlyCreatedDBObject(xr, true);
                }

                xr.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, value ?? ""));
            }
            catch
            {
                // Avoid command failure even if dictionary operations fail.
            }
        }
    }
}

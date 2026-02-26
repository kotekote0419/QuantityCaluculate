using System;
using System.Collections.Specialized;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;

namespace UFlowPlant3D.Services
{
    public static class QuantityKeyProp
    {
<<<<<<< HEAD
        private const string XREC_DICT = "UFLOW";
        private const string XREC_QTYID = "QTYID";
        private const string XREC_QTYKEY = "QTYKEY";

        // Plant側に作る（Project Setup）想定のプロパティ名候補
        private static readonly string[] QTYID_PROP_CANDIDATES = { "数量ID", "QuantityID", "QTY_ID" };
        private static readonly string[] QTYKEY_PROP_CANDIDATES = { "数量キー", "QtyKey", "QuantityKey", "QTY_KEY" };

        // ---------------------------
        // Entity(ObjectId) 用
        // ---------------------------
=======
        // DWG側に「無い場合に作る」保存先
        private const string XDICT_KEY = "UFLOW";
        private const string XREC_KEY = "QTYKEY";
        private const string XREC_LEGACY_KEY = "UFLOW:数量ID";
>>>>>>> origin/master

        /// <summary>
        /// 数量ID（数値文字列）を取得。Plant側に無ければXRecordから読む。
        /// </summary>
        public static string GetQuantityId(DataLinksManager dlm, Transaction tr, ObjectId oid)
        {
            var s = PlantProp.GetString(dlm, oid, QTYID_PROP_CANDIDATES);
            if (!string.IsNullOrWhiteSpace(s) && !LooksLikeLegacyKey(s)) return s.Trim();

<<<<<<< HEAD
            s = XRecordUtil.ReadString(tr, oid, XREC_DICT, XREC_QTYID);
            return s?.Trim() ?? "";
=======
            var fromDict = XRecordUtil.ReadString(tr, oid, XDICT_KEY, XREC_KEY);
            if (!string.IsNullOrWhiteSpace(fromDict)) return fromDict;

            var legacy = XRecordUtil.ReadLegacyString(tr, oid, XREC_LEGACY_KEY);
            if (!string.IsNullOrWhiteSpace(legacy))
            {
                XRecordUtil.WriteString(tr, oid, XDICT_KEY, XREC_KEY, legacy);
                return legacy;
            }

            return "";
>>>>>>> origin/master
        }

        /// <summary>
        /// 数量ID（数値文字列）をセット。Plant側に書けなければXRecordへ退避。
        /// </summary>
        public static void SetQuantityId(DataLinksManager dlm, Transaction tr, ObjectId oid, string qtyId)
        {
            qtyId ??= "";

            // まずPlantへ
            if (TryWritePlantQuantityId(dlm, oid, qtyId))
            {
                // 念のためXRecordにも退避（復旧・比較に使える）
                XRecordUtil.WriteString(tr, oid, XREC_DICT, XREC_QTYID, qtyId);
                return;
            }

<<<<<<< HEAD
            // Plantに書けないならXRecordで保持
            XRecordUtil.WriteString(tr, oid, XREC_DICT, XREC_QTYID, qtyId);
=======
            XRecordUtil.WriteString(tr, oid, XDICT_KEY, XREC_KEY, key);
>>>>>>> origin/master
        }

        /// <summary>
        /// 集計キー文字列（PIPE|...）を取得（XRecord優先）。
        /// </summary>
        public static string GetKey(DataLinksManager dlm, Transaction tr, ObjectId oid)
        {
            var s = XRecordUtil.ReadString(tr, oid, XREC_DICT, XREC_QTYKEY);
            if (!string.IsNullOrWhiteSpace(s)) return s.Trim();

            // 互換：旧仕様で「数量ID」にキーが入っていた場合
            s = PlantProp.GetString(dlm, oid, QTYID_PROP_CANDIDATES);
            if (LooksLikeLegacyKey(s)) return s.Trim();

            // もしProject Setupで「数量キー」を作っているならそこも読む
            s = PlantProp.GetString(dlm, oid, QTYKEY_PROP_CANDIDATES);
            if (!string.IsNullOrWhiteSpace(s)) return s.Trim();

            return "";
        }

        /// <summary>
        /// 集計キー文字列（PIPE|...）をXRecordへ保存。
        /// </summary>
        public static void SetKey(DataLinksManager dlm, Transaction tr, ObjectId oid, string key)
        {
            key ??= "";
            XRecordUtil.WriteString(tr, oid, XREC_DICT, XREC_QTYKEY, key);
        }

        private static bool LooksLikeLegacyKey(string s)
        {
            // 旧仕様： "PIPE|STW|架設|1100" のようなパイプ区切り
            return !string.IsNullOrWhiteSpace(s) && s.Contains("|");
        }

        // ---------------------------
        // RowId(FastenerRow) 用
        // ---------------------------

        public static string GetRowQuantityId(DataLinksManager dlm, int rowId)
        {
            var s = PlantProp.GetString(dlm, rowId, QTYID_PROP_CANDIDATES);
            if (!string.IsNullOrWhiteSpace(s) && !LooksLikeLegacyKey(s)) return s.Trim();
            return "";
        }

        public static void SetRowQuantityId(DataLinksManager dlm, int rowId, string qtyId)
        {
            qtyId ??= "";
            // RowId側はXRecordに退避できないので、書けるプロパティ名を順に試す
            TryWriteRowQuantityId(dlm, rowId, qtyId);
        }

        // ---------------------------
        // Backfill / Try-write helpers
        // ---------------------------

        /// <summary>
        /// Plantプロパティ「数量ID」に"だけ"書く（XRecordは触らない）。存在しない/書けないなら false。
        /// </summary>
        public static bool TryWritePlantQuantityId(DataLinksManager dlm, ObjectId oid, string qtyId)
        {
            qtyId ??= "";
            foreach (var name in QTYID_PROP_CANDIDATES)
            {
                if (TrySetPlantProp(dlm, oid, name, qtyId)) return true;
            }
            return false;
        }

        /// <summary>
        /// Plantプロパティ「数量キー」に"だけ"書く（任意）。存在しない/書けないなら false。
        /// </summary>
        public static bool TryWritePlantQuantityKey(DataLinksManager dlm, ObjectId oid, string key)
        {
            key ??= "";
            foreach (var name in QTYKEY_PROP_CANDIDATES)
            {
                if (TrySetPlantProp(dlm, oid, name, key)) return true;
            }
            return false;
        }

        /// <summary>
        /// rowId側の「数量ID」に"だけ"書く。存在しない/書けないなら false。
        /// </summary>
        public static bool TryWriteRowQuantityId(DataLinksManager dlm, int rowId, string qtyId)
        {
            qtyId ??= "";
            foreach (var name in QTYID_PROP_CANDIDATES)
            {
                if (TrySetRow(dlm, rowId, name, qtyId)) return true;
            }
            return false;
        }

        /// <summary>
        /// rowId側の「数量キー」に"だけ"書く（任意）。存在しない/書けないなら false。
        /// </summary>
        public static bool TryWriteRowQuantityKey(DataLinksManager dlm, int rowId, string key)
        {
            key ??= "";
            foreach (var name in QTYKEY_PROP_CANDIDATES)
            {
                if (TrySetRow(dlm, rowId, name, key)) return true;
            }
            return false;
        }

        // ---------------------------
        // Low-level SetProperties reflection
        // ---------------------------

        private static bool TrySetRow(DataLinksManager dlm, int rowId, string propName, string value)
        {
            try
            {
                var names = new StringCollection { propName };
                var vals = new StringCollection { value };
                if (InvokeSet(dlm, rowId, names, vals)) return true;

                if (InvokeSet(dlm, rowId, new[] { propName }, new[] { value })) return true;
                if (InvokeSet(dlm, rowId, new[] { propName }, new object[] { value })) return true;
            }
            catch { }
            return false;
        }

        private static bool TrySetPlantProp(DataLinksManager dlm, ObjectId oid, string propName, string value)
        {
            try
            {
                var names = new StringCollection { propName };
                var vals = new StringCollection { value };
                if (InvokeSet(dlm, oid, names, vals)) return true;

                if (InvokeSet(dlm, oid, new[] { propName }, new[] { value })) return true;
                if (InvokeSet(dlm, oid, new[] { propName }, new object[] { value })) return true;
            }
            catch { }
            return false;
        }

        private static bool InvokeSet(DataLinksManager dlm, object firstArg, object names, object vals)
        {
            try
            {
                foreach (var mi in dlm.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance))
                {
                    if (mi.Name != "SetProperties") continue;
                    var ps = mi.GetParameters();
                    if (ps.Length != 3) continue;

                    if (!ps[0].ParameterType.IsInstanceOfType(firstArg)) continue;
                    if (!ps[1].ParameterType.IsInstanceOfType(names)) continue;
                    if (!ps[2].ParameterType.IsInstanceOfType(vals)) continue;

                    mi.Invoke(dlm, new object[] { firstArg, names, vals });
                    return true;
                }
            }
            catch { }
            return false;
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
<<<<<<< HEAD
=======

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
>>>>>>> origin/master
    }
}

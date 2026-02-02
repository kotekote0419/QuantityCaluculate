using System;
using System.Collections.Specialized;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;

namespace UFlowPlant3D.Services
{
    public static class QuantityKeyProp
    {
        // DWG側に「無い場合に作る」保存先
        private const string XREC_KEY = "UFLOW:数量ID";

        /// <summary>
        /// 数量キー（文字列）を取得：
        /// 1) Plant3Dプロパティ（数量ID/QuantityID/QTY_ID）を優先
        /// 2) 無ければDWGのXRecordから
        /// </summary>
        public static string Get(DataLinksManager dlm, Transaction tr, ObjectId oid)
        {
            // Plant3D側
            var s = PlantProp.GetString(dlm, oid, "数量ID", "QuantityID", "QTY_ID");
            if (!string.IsNullOrWhiteSpace(s)) return s;

            // DWG側（XRecord）
            return XRecordUtil.ReadString(tr, oid, XREC_KEY) ?? "";
        }

        /// <summary>
        /// 数量キー（文字列）を設定：
        /// 1) Plant3Dプロパティに書ければ書く（SetPropertiesはStringCollection）
        /// 2) 書けなければDWGのXRecordへ保存（=「無い場合は作る」）
        /// </summary>
        public static void Set(DataLinksManager dlm, Transaction tr, ObjectId oid, string key)
        {
            key ??= "";

            // Plant3D側に書けるなら優先
            if (TrySetPlantProp(dlm, oid, "数量ID", key) ||
                TrySetPlantProp(dlm, oid, "QuantityID", key) ||
                TrySetPlantProp(dlm, oid, "QTY_ID", key))
            {
                return;
            }

            // 書けなかったらDWG側に保存（常に成功する）
            XRecordUtil.WriteString(tr, oid, XREC_KEY, key);
        }

        private static bool TrySetPlantProp(DataLinksManager dlm, ObjectId oid, string propName, string value)
        {
            try
            {
                int? rowId = PlantProp.TryGetRowIdPublic(dlm, oid);

                var names = new StringCollection { propName };
                var vals = new StringCollection { value };

                // シグネチャが存在する順に試す（あなたのDumpに一致）
                if (rowId.HasValue)
                {
                    // SetProperties(int, StringCollection, StringCollection)
                    if (InvokeSet(dlm, rowId.Value, names, vals)) return true;
                }

                // SetProperties(ObjectId, StringCollection, StringCollection)
                if (InvokeSet(dlm, oid, names, vals)) return true;

                // SetProperties(PpObjectId, StringCollection, StringCollection) も環境によってはある
                var ppoid = PlantProp.TryGetPpObjectId(dlm, oid);
                if (ppoid != null)
                {
                    if (InvokeSet(dlm, ppoid, names, vals)) return true;
                }
            }
            catch { }

            return false;
        }

        private static bool InvokeSet(DataLinksManager dlm, object firstArg, StringCollection names, StringCollection vals)
        {
            try
            {
                foreach (var mi in dlm.GetType().GetMethods())
                {
                    if (mi.Name != "SetProperties") continue;
                    var ps = mi.GetParameters();
                    if (ps.Length != 3) continue;

                    if (!ps[0].ParameterType.IsInstanceOfType(firstArg)) continue;
                    if (ps[1].ParameterType != typeof(StringCollection)) continue;
                    if (ps[2].ParameterType != typeof(StringCollection)) continue;

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
        public static string ReadString(Transaction tr, ObjectId entId, string key)
        {
            var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
            if (ent == null) return null;
            if (ent.ExtensionDictionary.IsNull) return null;

            var dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
            if (dict == null || !dict.Contains(key)) return null;

            var xr = tr.GetObject(dict.GetAt(key), OpenMode.ForRead) as Xrecord;
            if (xr?.Data == null) return null;

            var arr = xr.Data.AsArray();
            if (arr == null || arr.Length == 0) return null;

            return arr[0].Value as string;
        }

        public static void WriteString(Transaction tr, ObjectId entId, string key, string value)
        {
            var ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
            if (ent == null) return;

            if (ent.ExtensionDictionary.IsNull)
                ent.CreateExtensionDictionary();

            var dict = tr.GetObject(ent.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;
            if (dict == null) return;

            Xrecord xr;
            if (dict.Contains(key))
            {
                xr = tr.GetObject(dict.GetAt(key), OpenMode.ForWrite) as Xrecord;
            }
            else
            {
                xr = new Xrecord();
                dict.SetAt(key, xr);
                tr.AddNewlyCreatedDBObject(xr, true);
            }

            xr.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, value ?? ""));
        }
    }
}

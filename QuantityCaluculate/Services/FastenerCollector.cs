using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;

namespace UFlowPlant3D.Services
{
    /// <summary>
    /// PnPDatabase の "Fasteners" テーブルから Fastener の ObjectId を回収する。
    /// 参照DLL差を吸収するため、Autodesk.ProcessPower.PnPDatabase 名前空間には依存しない（反射）。
    /// また、FindAcPpObjectIds の戻り値が ObjectId[] / ObjectIdCollection 等でも吸収する。
    /// </summary>
    public static class FastenerCollector
    {
        public static List<ObjectId> CollectFastenerObjectIds(DataLinksManager dlm)
        {
            var ids = new List<ObjectId>();
            try
            {
                var proj = PlantApplication.CurrentProject;
                if (proj == null) return ids;

                var pipingPart = proj.ProjectParts["Piping"];
                var pDlm = pipingPart?.DataLinksManager;
                if (pDlm == null) return ids;

                // pDlm.PnPDatabase を反射で取得
                var pnpDb = Reflect.GetProp(pDlm, "PnPDatabase");
                if (pnpDb == null) return ids;

                // pnpDb.Tables を取得（テーブルコレクション）
                var tables = Reflect.GetProp(pnpDb, "Tables");
                if (tables == null) return ids;

                // tables["Fasteners"] を取得（indexer/get_Item）
                var fastTable = Reflect.GetIndexer(tables, "Fasteners");
                if (fastTable == null) return ids;

                // fastTable.Rows を列挙
                var rowsObj = Reflect.GetProp(fastTable, "Rows");
                if (rowsObj is not IEnumerable rows) return ids;

                foreach (var row in rows)
                {
                    // row.RowId（int）を取得
                    var rowIdObj = Reflect.GetProp(row, "RowId");
                    if (!Reflect.TryToInt(rowIdObj, out int rowId)) continue;

                    // rowId -> ObjectId(s)
                    object oidsObj = null;
                    try { oidsObj = dlm.FindAcPpObjectIds(rowId); } catch { /* ignore */ }

                    foreach (var oid in ObjectIdUtil.ToObjectIds(oidsObj))
                    {
                        if (oid.IsNull || oid.IsErased) continue;
                        ids.Add(oid);
                    }
                }
            }
            catch
            {
                // 必要ならログ出し（ここは静かに落とす）
            }

            // 重複排除して返す
            return new List<ObjectId>(new HashSet<ObjectId>(ids));
        }
    }

    internal static class ObjectIdUtil
    {
        /// <summary>
        /// FindAcPpObjectIds の戻り値が ObjectId[] / ObjectIdCollection / IEnumerable のいずれでも吸収して列挙する。
        /// </summary>
        public static IEnumerable<ObjectId> ToObjectIds(object oidsObj)
        {
            if (oidsObj == null) yield break;

            // 1) 配列
            if (oidsObj is ObjectId[] arr)
            {
                foreach (var oid in arr) yield return oid;
                yield break;
            }

            // 2) IEnumerable（ObjectIdCollection 等）
            if (oidsObj is IEnumerable e)
            {
                foreach (var x in e)
                {
                    if (x is ObjectId oid) yield return oid;
                }
                yield break;
            }

            // 3) それ以外は諦める
            yield break;
        }
    }

    internal static class Reflect
    {
        public static object GetProp(object obj, string name)
        {
            if (obj == null) return null;
            try
            {
                var pi = obj.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                return pi?.GetValue(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// indexer this[string] / get_Item(string) の両方を試して取得
        /// </summary>
        public static object GetIndexer(object obj, string key)
        {
            if (obj == null) return null;

            try
            {
                // C# indexer: Item(string)
                var pi = obj.GetType().GetProperty("Item", BindingFlags.Public | BindingFlags.Instance);
                if (pi != null)
                {
                    var idx = pi.GetIndexParameters();
                    if (idx.Length == 1 && idx[0].ParameterType == typeof(string))
                    {
                        try { return pi.GetValue(obj, new object[] { key }); } catch { /* ignore */ }
                    }
                }
            }
            catch { /* ignore */ }

            try
            {
                // get_Item(string)
                var mi = obj.GetType().GetMethod("get_Item", BindingFlags.Public | BindingFlags.Instance, null,
                    new[] { typeof(string) }, null);
                if (mi != null)
                {
                    try { return mi.Invoke(obj, new object[] { key }); } catch { /* ignore */ }
                }
            }
            catch { /* ignore */ }

            return null;
        }

        public static bool TryToInt(object v, out int i)
        {
            i = 0;
            if (v == null) return false;
            if (v is int ii) { i = ii; return true; }
            return int.TryParse(v.ToString(), out i);
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;

using UFlowPlant3D.Services;

namespace UFlowPlant3D.Commands
{
    public class DebugGasketPropsCommands
    {
        [CommandMethod("UFLOW_DEBUG_DUMP_GASKET_LINK_PROPS")]
        public void UFLOW_DEBUG_DUMP_GASKET_LINK_PROPS()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var dlm = PlantApplication.CurrentProject?.ProjectParts["Piping"]?.DataLinksManager;
            if (dlm == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] DataLinksManagerが取得できません。");
                return;
            }

            // 1) pick
            var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick gasket (Connector/P3dConnector): ");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using var tr = db.TransactionManager.StartTransaction();
            var oid = per.ObjectId;
            var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
            if (ent == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] Selected object is not Entity.");
                return;
            }

            string handle = "";
            try { handle = ent.Handle.ToString(); } catch { }

            int rowId = TryFindRowId(dlm, oid);

            ed.WriteMessage($"\n[UFLOW][DBG] Picked: Handle={handle} Type={ent.GetType().FullName} RowId={rowId}");

            // 2) dump DLM props by ObjectId
            ed.WriteMessage("\n[UFLOW][DBG] ---- DLM GetAllProperties(ObjectId) ----");
            DumpAllProps(ed, InvokeProps(dlm, oid), 500);

            // 3) dump DLM props by RowId
            if (rowId > 0)
            {
                ed.WriteMessage("\n[UFLOW][DBG] ---- DLM GetAllProperties(RowId) ----");
                DumpAllProps(ed, InvokeProps(dlm, rowId), 500);
            }

            string guid = "";
            if (rowId > 0) guid = (PlantProp.GetString(dlm, rowId, "PnPGuid", "RowGuid", "GUID") ?? "").Trim();
            if (string.IsNullOrWhiteSpace(guid))
                guid = (PlantProp.GetString(dlm, oid, "PnPGuid", "RowGuid", "GUID") ?? "").Trim();

            if (!string.IsNullOrWhiteSpace(guid))
                ed.WriteMessage($"\n[UFLOW][DBG] Picked PnPGuid='{guid}'");

            // 4) scan PnPDatabase tables for linkage
            ed.WriteMessage("\n[UFLOW][DBG] ---- Scan PnPDB tables: Fasteners / Gasket / BoltSet ----");
            ScanTables(ed, dlm, rowId, guid, 15);

            tr.Commit();
        }

        // ------------------------
        // RowId getter
        // ------------------------
        private static int TryFindRowId(DataLinksManager dlm, ObjectId oid)
        {
            string[] names = { "FindAcPpRowId", "FindRowId", "GetRowId" };
            foreach (var n in names)
            {
                try
                {
                    var mi = dlm.GetType().GetMethod(n, BindingFlags.Public | BindingFlags.Instance,
                        null, new[] { typeof(ObjectId) }, null);
                    if (mi == null) continue;
                    var v = mi.Invoke(dlm, new object[] { oid });
                    if (v is int i && i > 0) return i;
                }
                catch { }
            }
            return 0;
        }

        // ------------------------
        // Invoke GetAllProperties / GetProperties
        // ------------------------
        private static object InvokeProps(DataLinksManager dlm, ObjectId oid)
        {
            var t = dlm.GetType();
            try
            {
                var mi = t.GetMethod("GetAllProperties", BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(ObjectId), typeof(bool) }, null);
                if (mi != null) return mi.Invoke(dlm, new object[] { oid, true });
            }
            catch { }
            try
            {
                var mi = t.GetMethod("GetProperties", BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(ObjectId), typeof(bool) }, null);
                if (mi != null) return mi.Invoke(dlm, new object[] { oid, true });
            }
            catch { }
            return null;
        }

        private static object InvokeProps(DataLinksManager dlm, int rowId)
        {
            var t = dlm.GetType();
            try
            {
                var mi = t.GetMethod("GetAllProperties", BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(int), typeof(bool) }, null);
                if (mi != null) return mi.Invoke(dlm, new object[] { rowId, true });
            }
            catch { }
            try
            {
                var mi = t.GetMethod("GetProperties", BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(int), typeof(bool) }, null);
                if (mi != null) return mi.Invoke(dlm, new object[] { rowId, true });
            }
            catch { }
            return null;
        }

        // ------------------------
        // Dump all props object
        // ------------------------
        private static void DumpAllProps(Editor ed, object props, int maxLines)
        {
            if (props == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] (props=null)");
                return;
            }

            if (props is IDictionary dict)
            {
                int n = 0;
                foreach (DictionaryEntry de in dict)
                {
                    if (n++ >= maxLines) break;
                    ed.WriteMessage($"\n[UFLOW][DBG]  {de.Key} = {de.Value}");
                }
                ed.WriteMessage($"\n[UFLOW][DBG]  (count~={dict.Count})");
                return;
            }

            if (props is IEnumerable en)
            {
                int n = 0;
                foreach (var item in en)
                {
                    if (item == null) continue;
                    if (n++ >= maxLines) break;

                    string name = GetPropString(item, "Name")
                               ?? GetPropString(item, "Key")
                               ?? GetPropString(item, "PropertyName")
                               ?? GetPropString(item, "DisplayName")
                               ?? "";
                    object val = GetPropObj(item, "Value")
                              ?? GetPropObj(item, "PropValue")
                              ?? GetPropObj(item, "PropertyValue")
                              ?? item.ToString();

                    ed.WriteMessage($"\n[UFLOW][DBG]  {name} = {val}");
                }
                ed.WriteMessage($"\n[UFLOW][DBG]  (enumerated~={n})");
                return;
            }

            ed.WriteMessage($"\n[UFLOW][DBG] propsType={props.GetType().FullName} ToString={props}");
        }

        private static object GetPropObj(object obj, string name)
        {
            try
            {
                var pi = obj.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                return pi?.GetValue(obj);
            }
            catch { return null; }
        }
        private static string GetPropString(object obj, string name) => GetPropObj(obj, name)?.ToString();

        // ------------------------
        // Scan PnPDB tables
        // ------------------------
        private static void ScanTables(Editor ed, DataLinksManager dlm, int pickedRowId, string pickedGuid, int maxShowHits)
        {
            object pnpDb = Reflect.GetProp(dlm, "PnPDatabase") ?? Reflect.Invoke0(dlm, "GetPnPDatabase");
            if (pnpDb == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] PnPDatabase not available.");
                return;
            }

            object tables = Reflect.GetProp(pnpDb, "Tables") ?? Reflect.Invoke0(pnpDb, "GetTables");
            if (tables == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] PnPDatabase.Tables not available.");
                return;
            }

            string[] scan = { "Fasteners", "Gasket", "BoltSet" };
            foreach (var tname in scan)
            {
                object table = Reflect.GetIndexer(tables, tname) ?? Reflect.Invoke1(pnpDb, "GetTable", tname);
                if (table == null)
                {
                    ed.WriteMessage($"\n[UFLOW][DBG] Table '{tname}' not found.");
                    continue;
                }

                int hits = 0, shown = 0;
                foreach (var row in EnumeratePnPTableRows(table))
                {
                    if (row == null) continue;

                    if (RowContains(row, pickedRowId, pickedGuid))
                    {
                        hits++;
                        if (shown < maxShowHits)
                        {
                            shown++;
                            int rid = TryGetIntProp(row, "RowId");
                            string cls = TryGetStringProp(row, "ClassName") ?? "";
                            ed.WriteMessage($"\n[UFLOW][DBG] HIT table='{tname}' rowId={rid} class='{cls}'");
                            DumpRowProps(ed, row, 50);
                        }
                    }
                }

                ed.WriteMessage($"\n[UFLOW][DBG] Table '{tname}' hits={hits}");
            }
        }

        private static IEnumerable EnumeratePnPTableRows(object table)
        {
            // PnPTable.Select(\"1=1\") を優先
            IEnumerable selected = TrySelectAllRows(table);
            if (selected != null)
            {
                foreach (var r in selected) yield return r;
                yield break;
            }

            object rowsObj = Reflect.GetProp(table, "Rows") ?? Reflect.Invoke0(table, "GetRows");
            if (rowsObj is IEnumerable en)
            {
                foreach (var r in en) yield return r;
            }
        }

        private static IEnumerable TrySelectAllRows(object table)
        {
            if (table == null) return null;
            var t = table.GetType();
            object ret = null;
            try
            {
                var mi = t.GetMethod("Select", BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(string) }, null);
                if (mi != null) ret = mi.Invoke(table, new object[] { "1=1" });
            }
            catch { }
            if (ret is IEnumerable e) return e;
            return null;
        }

        private static bool RowContains(object row, int id, string guid)
        {
            if (row == null) return false;
            foreach (var p in row.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (p.GetIndexParameters().Length > 0) continue;
                object v = null;
                try { v = p.GetValue(row); } catch { }
                if (v == null) continue;

                if (id > 0 && int.TryParse(v.ToString(), out var vi) && vi == id) return true;
                if (!string.IsNullOrWhiteSpace(guid) && v.ToString().IndexOf(guid, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }

        private static void DumpRowProps(Editor ed, object row, int max)
        {
            int n = 0;
            foreach (var p in row.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (n++ >= max) break;
                if (p.GetIndexParameters().Length > 0) continue;
                object v = null;
                try { v = p.GetValue(row); } catch { }
                ed.WriteMessage($"\n[UFLOW][DBG]   {p.Name} = {v}");
            }
        }

        private static int TryGetIntProp(object obj, string name)
        {
            try
            {
                var pi = obj.GetType().GetProperty(name);
                var v = pi?.GetValue(obj);
                if (v == null) return 0;
                if (v is int i) return i;
                int.TryParse(v.ToString(), out var ii);
                return ii;
            }
            catch { return 0; }
        }
        private static string TryGetStringProp(object obj, string name)
        {
            try
            {
                var pi = obj.GetType().GetProperty(name);
                return pi?.GetValue(obj)?.ToString();
            }
            catch { return null; }
        }
    }
}

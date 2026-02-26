using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;

namespace UFlow
{
    public class DebugPnPDataLinksCommands
    {
        [CommandMethod("UFLOW_DEBUG_DUMP_PNPDATALINKS_ROW")]
        public void UFLOW_DEBUG_DUMP_PNPDATALINKS_ROW()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var dlm = SafeGetDlm(db, ed);
            if (dlm == null) return;

            var pnpDb = SafeGetPnPDatabase(dlm, ed);
            if (pnpDb == null) return;

            var p = ed.GetInteger("\n[UFLOW][DBG] Input PnPDataLinks RowId (e.g. 32275): ");
            if (p.Status != PromptStatus.OK) return;

            int linkRowId = p.Value;

            var table = pnpDb.Tables["PnPDataLinks"];
            if (table == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] Table 'PnPDataLinks' not found.");
                return;
            }

            var colNames = GetColumnNames(table, out bool fallback);
            ed.WriteMessage($"\n[UFLOW][DBG] PnPDataLinks columns={colNames.Count} (fallback={fallback})");

            if (!TrySelectAllRowsFromPnPTable(table, out var rows))
            {
                ed.WriteMessage("\n[UFLOW][DBG] PnPDataLinks Select failed.");
                return;
            }

            object hit = null;
            foreach (var r in rows)
            {
                int rid = TryGetIntProp(r, "RowId");
                if (rid == linkRowId) { hit = r; break; }
            }

            if (hit == null)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] Link RowId={linkRowId} not found in PnPDataLinks.");
                return;
            }

            ed.WriteMessage($"\n[UFLOW][DBG] LinkRow type={hit.GetType().FullName}");
            ed.WriteMessage($"\n[UFLOW][DBG] ---- Dump all columns for RowId={linkRowId} ----");

            foreach (var c in colNames)
            {
                string v = TryGetRowValueAsString(hit, c);
                if (v == null) v = "";
                if (v.Length > 200) v = v.Substring(0, 200) + "...";
                ed.WriteMessage($"\n[UFLOW][DBG]  {c} = {v}");
            }
        }

        [CommandMethod("UFLOW_DEBUG_FIND_DATALINKS_FOR_PICKED_CONNECTOR")]
        public void UFLOW_DEBUG_FIND_DATALINKS_FOR_PICKED_CONNECTOR()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var dlm = SafeGetDlm(db, ed);
            if (dlm == null) return;

            // pick connector-like entity
            var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick Connector (gasket-like): ");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            int connRowId;
            try { connRowId = dlm.FindAcPpRowId(per.ObjectId); }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] FindAcPpRowId failed: {ex.Message}");
                return;
            }

            ed.WriteMessage($"\n[UFLOW][DBG] Connector RowId={connRowId}");

            var pnpDb = SafeGetPnPDatabase(dlm, ed);
            if (pnpDb == null) return;

            var table = pnpDb.Tables["PnPDataLinks"];
            if (table == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] Table 'PnPDataLinks' not found.");
                return;
            }

            var colNames = GetColumnNames(table, out bool fallback);
            ed.WriteMessage($"\n[UFLOW][DBG] PnPDataLinks columns={colNames.Count} (fallback={fallback})");

            if (!TrySelectAllRowsFromPnPTable(table, out var rows))
            {
                ed.WriteMessage("\n[UFLOW][DBG] PnPDataLinks Select failed.");
                return;
            }

            string idStr = connRowId.ToString();
            int hits = 0;

            foreach (var r in rows)
            {
                // 「どの列に32278が入ってるか」を列名付きで拾う
                var hitCols = new List<string>();
                foreach (var c in colNames)
                {
                    string v = TryGetRowValueAsString(r, c);
                    if (string.IsNullOrWhiteSpace(v)) continue;
                    if (v.IndexOf(idStr, StringComparison.OrdinalIgnoreCase) >= 0)
                        hitCols.Add($"{c}={Trim(v)}");
                }

                if (hitCols.Count > 0)
                {
                    int rid = TryGetIntProp(r, "RowId");
                    ed.WriteMessage($"\n[UFLOW][DBG] HIT LinkRowId={rid} cols={hitCols.Count}");
                    foreach (var hc in hitCols.Take(12)) ed.WriteMessage($"\n[UFLOW][DBG]   {hc}");
                    if (hitCols.Count > 12) ed.WriteMessage($"\n[UFLOW][DBG]   ... (more {hitCols.Count - 12})");
                    hits++;
                    if (hits >= 30)
                    {
                        ed.WriteMessage("\n[UFLOW][DBG] stop (hits>=30)");
                        break;
                    }
                }
            }

            ed.WriteMessage($"\n[UFLOW][DBG] Done. datalinksHits={hits}");
        }

        // -------- shared helpers --------

        private static string Trim(string s)
        {
            if (s == null) return "";
            s = s.Trim();
            return s.Length > 120 ? s.Substring(0, 120) + "..." : s;
        }

        private static DataLinksManager SafeGetDlm(Database db, Editor ed)
        {
            try { return DataLinksManager.GetManager(db); }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] DLM get failed: {ex.Message}");
                return null;
            }
        }

        private static PnPDatabase SafeGetPnPDatabase(DataLinksManager dlm, Editor ed)
        {
            try { return dlm.GetPnPDatabase(); }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] GetPnPDatabase failed: {ex.Message}");
                return null;
            }
        }

        private static bool TrySelectAllRowsFromPnPTable(PnPTable table, out IEnumerable rows)
        {
            rows = null;
            if (table == null) return false;

            object ret = null;
            var t = table.GetType();

            try
            {
                var mi = t.GetMethod("Select", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                                     null, new[] { typeof(string) }, null);
                if (mi != null) ret = mi.Invoke(table, new object[] { "1=1" });
            }
            catch { }

            if (ret == null)
            {
                try
                {
                    var mi0 = t.GetMethod("Select", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                                          null, Type.EmptyTypes, null);
                    if (mi0 != null) ret = mi0.Invoke(table, null);
                }
                catch { }
            }

            if (ret == null) return false;

            if (ret is IEnumerable e)
            {
                rows = e;
                return true;
            }

            rows = new object[] { ret };
            return true;
        }

        private static List<string> GetColumnNames(PnPTable table, out bool fallback)
        {
            fallback = false;
            var names = new List<string>();

            try
            {
                var pi = table.GetType().GetProperty("Columns");
                var cols = pi?.GetValue(table, null) as IEnumerable;
                if (cols != null)
                {
                    foreach (var c in cols)
                    {
                        var n = c.GetType().GetProperty("Name")?.GetValue(c, null)?.ToString()
                             ?? c.GetType().GetProperty("ColumnName")?.GetValue(c, null)?.ToString();
                        if (!string.IsNullOrWhiteSpace(n)) names.Add(n);
                    }
                }
            }
            catch { }

            if (names.Count == 0)
            {
                fallback = true;
                names.AddRange(new[] { "RowId" });
            }

            return names.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string TryGetRowValueAsString(object row, string colName)
        {
            try
            {
                var prop = row.GetType().GetProperty("Item", new[] { typeof(string) });
                if (prop != null)
                {
                    var v = prop.GetValue(row, new object[] { colName });
                    return v?.ToString() ?? "";
                }
            }
            catch { }
            return "";
        }

        private static int TryGetIntProp(object obj, string name)
        {
            try
            {
                var pi = obj.GetType().GetProperty(name);
                var v = pi?.GetValue(obj, null);
                if (v == null) return -1;
                if (v is int i) return i;
                if (int.TryParse(v.ToString(), out var ii)) return ii;
            }
            catch { }
            return -1;
        }
    }
}

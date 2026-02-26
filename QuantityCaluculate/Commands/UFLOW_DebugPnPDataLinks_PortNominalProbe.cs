// UFLOW_DebugPnPDataLinks_PortNominalProbe.cs
// Debug command to explore "Approach A":
// Try to locate per-port NominalDiameter (ND) by traversing PnPDataLinks bundle rows,
// then reading DLM properties of the referenced RowId (RowClassName) rows.
//
// Command:
//   UFLOW_DEBUG_PNPDATALINKS_PORTNOMINAL
//
// Target:
// - For a picked entity (e.g., Tee), find PnPDataLinks rows that share the same DwgHandleLow/High (+DwgId if present)
// - For each bundle row, resolve (RowId, RowClassName, DwgSubIndex)
// - For each referenced RowId, call dlm.GetAllProperties(rowId,true) and log occurrences of:
//     PortName, NominalDiameter, NominalUnit
//   (in list order, without collapsing duplicate keys)
//
// Notes:
// - Plant 3D 2026: PnPDatabase access may be behind non-public members.
// - This file avoids dynamic; uses reflection to disambiguate overloads.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.DataLinks;

namespace UFLOW
{
    public class UFLOW_DebugPnPDataLinks_PortNominalProbe
    {
        [CommandMethod("UFLOW_DEBUG_PNPDATALINKS_PORTNOMINAL")]
        public void DebugPnPDataLinksPortNominal()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick entity to probe PnPDataLinks (port ND): ");
            peo.AllowNone = false;

            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                Log(ed, "[UFLOW][DBG] Cancelled.");
                return;
            }

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead, false) as Entity;
                if (ent == null)
                {
                    Log(ed, "[UFLOW][DBG] Not an Entity.");
                    return;
                }

                Log(ed, $"[UFLOW][DBG] Picked: Handle={ent.Handle} Type={ent.GetType().FullName}");

                var dlm = TryGetDataLinksManager(ed);
                if (dlm == null)
                {
                    Log(ed, "[UFLOW][DBG] DataLinksManager not available.");
                    return;
                }

                int baseRowId = -1;
                try { baseRowId = InvokeFindAcPpRowId_ObjectId(dlm, per.ObjectId); } catch { /* ignore */ }
                Log(ed, $"[UFLOW][DBG] BaseObject rowId={baseRowId}");
                if (baseRowId <= 0)
                {
                    Log(ed, "[UFLOW][DBG] Base rowId <= 0. Abort.");
                    return;
                }

                object pnpDb = TryGetPnPDatabaseFromDlm(dlm, ed);
                if (pnpDb == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDatabase not available from DataLinksManager.");
                    return;
                }

                object links = TryGetPnPTable(pnpDb, "PnPDataLinks");
                if (links == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDataLinks table not found in PnPDatabase.");
                    return;
                }

                // ---- 1) find "baseRows" for this entity ----
                IEnumerable baseRows = null;
                string usedWhere = null;

                foreach (var w in new[]
                {
                    $"RowId={baseRowId}",
                    $"PnPID={baseRowId}",
                    $"PnPRowId={baseRowId}",
                    $"PnPId={baseRowId}"
                })
                {
                    var rr = TrySelectRows(links, w);
                    if (rr == null) continue;

                    var en = rr.GetEnumerator();
                    if (en != null && en.MoveNext())
                    {
                        baseRows = rr;
                        usedWhere = w;
                        break;
                    }
                }

                if (baseRows == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDataLinks: no baseRows found for this rowId (tried RowId/PnPID/PnPRowId).");
                    return;
                }

                Log(ed, $"[UFLOW][DBG] PnPDataLinks: baseRows found by where='{usedWhere}'");

                object first = null;
                foreach (var r in baseRows) { first = r; break; }
                if (first == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDataLinks: baseRows empty.");
                    return;
                }

                long dwgId = TryToInt64(TryGetRowField(first, "DwgId"));
                long low = TryToInt64(TryGetRowField(first, "DwgHandleLow"));
                long high = TryToInt64(TryGetRowField(first, "DwgHandleHigh"));

                Log(ed, $"[UFLOW][DBG] PnPDataLinks: DwgId={dwgId} DwgHandleLow={low} DwgHandleHigh={high}");

                if (low == 0 && high == 0)
                {
                    Log(ed, "[UFLOW][DBG] PnPDataLinks: handle fields are zero/not found. Dumping first row fields (first 20 keys)...");
                    DumpRowFields(ed, first, maxKeys: 20);
                    return;
                }

                // ---- 2) bundle query by handle (+dwgId if available) ----
                string whereHandle = (high != 0)
                    ? $"DwgHandleLow={low} AND DwgHandleHigh={high}"
                    : $"DwgHandleLow={low}";

                if (dwgId != 0)
                {
                    whereHandle = $"({whereHandle}) AND DwgId={dwgId}";
                }

                var bundle = TrySelectRows(links, whereHandle);
                if (bundle == null)
                {
                    Log(ed, $"[UFLOW][DBG] PnPDataLinks: Select returned null. where='{whereHandle}'");
                    return;
                }

                // ---- 3) iterate bundle and try to read per-row props ----
                Log(ed, "[UFLOW][DBG] ---- PnPDataLinks bundle rows ----");

                int idx = 0;
                foreach (var r in bundle)
                {
                    idx++;

                    long rowId = TryToInt64(TryGetRowField(r, "RowId"));
                    long pnpid = TryToInt64(TryGetRowField(r, "PnPID"));
                    long sub = TryToInt64(TryGetRowField(r, "DwgSubIndex"));
                    string rowClass = TryToString(TryGetRowField(r, "RowClassName"));

                    if (pnpid == 0) pnpid = rowId;

                    Log(ed, $"[UFLOW][DBG] [LinkRow {idx}] PnPID={pnpid} RowId={rowId} RowClassName='{rowClass}' DwgSubIndex={sub}");

                    if (rowId <= 0) continue;

                    List<KeyValuePair<string, string>> props = null;
                    try
                    {
                        props = InvokeGetAllProperties(dlm, (int)rowId, true);
                    }
                    catch (System.Exception ex)
                    {
                        Log(ed, $"[UFLOW][DBG]   GetAllProperties(rowId={rowId}) failed: {ex.Message}");
                        continue;
                    }

                    if (props == null || props.Count == 0)
                    {
                        Log(ed, $"[UFLOW][DBG]   GetAllProperties(rowId={rowId}) => (empty)");
                        continue;
                    }

                    string ln = FindFirst(props, "LineNumberTag");
                    string sz = FindFirst(props, "Size");
                    string desc = FindFirst(props, "PartFamilyLongDesc");
                    string it = FindFirst(props, "InstallType");
                    if (!string.IsNullOrWhiteSpace(ln) || !string.IsNullOrWhiteSpace(sz) || !string.IsNullOrWhiteSpace(desc) || !string.IsNullOrWhiteSpace(it))
                    {
                        Log(ed, $"[UFLOW][DBG]   KeyProps: LineNumberTag='{ln}' Size='{sz}' InstallType='{it}' PartFamilyLongDesc='{desc}'");
                    }

                    var occ = props.Where(kv =>
                            IsKey(kv.Key, "PortName") ||
                            IsKey(kv.Key, "NominalDiameter") ||
                            IsKey(kv.Key, "NominalUnit"))
                        .ToList();

                    if (occ.Count == 0)
                    {
                        var contains = props.Where(kv =>
                                ContainsKey(kv.Key, "nominal") ||
                                ContainsKey(kv.Key, "port"))
                            .Take(12)
                            .ToList();

                        Log(ed, $"[UFLOW][DBG]   Port/Nominal occurrences: (none). contains(port/nominal) sample count={contains.Count}");
                        foreach (var kv in contains)
                            Log(ed, $"[UFLOW][DBG]    {kv.Key} = {kv.Value}");
                        continue;
                    }

                    Log(ed, $"[UFLOW][DBG]   ---- PortName/NominalDiameter/NominalUnit occurrences (count={occ.Count}) ----");
                    foreach (var kv in occ)
                        Log(ed, $"[UFLOW][DBG]    {kv.Key} = {kv.Value}");
                    Log(ed, "[UFLOW][DBG]   ---- occurrences end ----");
                }

                Log(ed, "[UFLOW][DBG] ---- bundle end ----");

                tr.Commit();
            }
        }

        // ---------------- helpers ----------------

        private static void Log(Editor ed, string msg)
        {
            if (ed == null) return;
            ed.WriteMessage("\n" + msg);
        }

        private static DataLinksManager TryGetDataLinksManager(Editor ed)
        {
            try
            {
                var proj = PlantApplication.CurrentProject;
                if (proj == null) return null;

                var pp = proj.ProjectParts;
                if (pp == null) return null;

                var idx = pp.GetType().GetProperty("Item", new Type[] { typeof(string) });
                if (idx == null) return null;

                var piping = idx.GetValue(pp, new object[] { "Piping" });
                if (piping == null) return null;

                var p = piping.GetType().GetProperty("DataLinksManager", BindingFlags.Public | BindingFlags.Instance);
                if (p == null) return null;

                var dlm = p.GetValue(piping, null) as DataLinksManager;
                return dlm;
            }
            catch (System.Exception ex)
            {
                Log(ed, "[UFLOW][DBG] TryGetDataLinksManager failed: " + ex.Message);
                return null;
            }
        }

        private static object TryGetPnPDatabaseFromDlm(DataLinksManager dlm, Editor ed)
        {
            try
            {
                var t = dlm.GetType();
                var mi = t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                          .FirstOrDefault(m => m.Name == "GetPnPDatabase" && m.GetParameters().Length == 0);
                if (mi != null)
                {
                    var v = mi.Invoke(dlm, null);
                    if (v != null)
                    {
                        Log(ed, $"[UFLOW][DBG] Candidate via DataLinksManager.GetPnPDatabase() type={v.GetType().FullName}");
                        return v;
                    }
                }
            }
            catch (System.Exception ex)
            {
                Log(ed, "[UFLOW][DBG] GetPnPDatabase() invoke failed: " + ex.Message);
            }

            try
            {
                var flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
                var t = dlm.GetType();

                foreach (var p in t.GetProperties(flags))
                {
                    if (!p.CanRead) continue;
                    object v = null;
                    try { v = p.GetValue(dlm, null); } catch { continue; }
                    if (v == null) continue;
                    var n = v.GetType().FullName ?? "";
                    if (n.IndexOf("PnPDatabase", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        n.IndexOf("Mode", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        Log(ed, $"[UFLOW][DBG] Candidate via dlm.{p.Name} type={n}");
                        return v;
                    }
                }

                foreach (var f in t.GetFields(flags))
                {
                    object v = null;
                    try { v = f.GetValue(dlm); } catch { continue; }
                    if (v == null) continue;
                    var n = v.GetType().FullName ?? "";
                    if (n.IndexOf("PnPDatabase", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        n.IndexOf("Mode", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        Log(ed, $"[UFLOW][DBG] Candidate via dlm.{f.Name} (field) type={n}");
                        return v;
                    }
                }
            }
            catch { }

            return null;
        }

        private static object TryGetPnPTable(object pnpDb, string tableName)
        {
            try
            {
                var t = pnpDb.GetType();
                var pTables = t.GetProperty("Tables", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (pTables == null) return null;

                var tables = pTables.GetValue(pnpDb, null);
                if (tables == null) return null;

                var idx = tables.GetType().GetProperty("Item", new Type[] { typeof(string) });
                if (idx == null) return null;

                return idx.GetValue(tables, new object[] { tableName });
            }
            catch
            {
                return null;
            }
        }

        private static IEnumerable TrySelectRows(object pnpTable, string where)
        {
            try
            {
                var mi = pnpTable.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance)
                                 .FirstOrDefault(m => m.Name == "Select" && m.GetParameters().Length == 1 && m.GetParameters()[0].ParameterType == typeof(string));
                if (mi == null) return null;

                var v = mi.Invoke(pnpTable, new object[] { where });
                return v as IEnumerable;
            }
            catch
            {
                return null;
            }
        }

        private static object TryGetRowField(object row, string fieldName)
        {
            if (row == null || string.IsNullOrWhiteSpace(fieldName)) return null;

            try
            {
                var t = row.GetType();
                var idx = t.GetProperty("Item", new Type[] { typeof(string) });
                if (idx != null)
                {
                    return idx.GetValue(row, new object[] { fieldName });
                }
            }
            catch { }

            try
            {
                var p = row.GetType().GetProperty(fieldName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (p != null) return p.GetValue(row, null);
            }
            catch { }

            return null;
        }

        private static void DumpRowFields(Editor ed, object row, int maxKeys)
        {
            if (row == null) return;

            try
            {
                var colsProp = row.GetType().GetProperty("Columns", BindingFlags.Public | BindingFlags.Instance);
                var cols = colsProp?.GetValue(row, null) as IEnumerable;
                if (cols == null) return;

                int i = 0;
                foreach (var c in cols)
                {
                    if (++i > maxKeys) break;
                    string name = TryToString(c);
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    var v = TryGetRowField(row, name);
                    Log(ed, $"[UFLOW][DBG]  {name} = {TryToString(v)}");
                }
            }
            catch { }
        }

        private static int InvokeFindAcPpRowId_ObjectId(DataLinksManager dlm, ObjectId oid)
        {
            var mi = dlm.GetType().GetMethod("FindAcPpRowId", new Type[] { typeof(ObjectId) });
            if (mi == null) throw new MissingMethodException("FindAcPpRowId(ObjectId) not found");
            return (int)mi.Invoke(dlm, new object[] { oid });
        }

        private static List<KeyValuePair<string, string>> InvokeGetAllProperties(DataLinksManager dlm, int rowId, bool inherit)
        {
            var mi = dlm.GetType().GetMethod("GetAllProperties", new Type[] { typeof(int), typeof(bool) });
            if (mi == null) throw new MissingMethodException("GetAllProperties(int,bool) not found");
            return mi.Invoke(dlm, new object[] { rowId, inherit }) as List<KeyValuePair<string, string>>;
        }

        private static bool IsKey(string k, string target)
        {
            if (k == null) return false;
            return string.Equals(k.Trim(), target, StringComparison.OrdinalIgnoreCase);
        }

        private static bool ContainsKey(string k, string part)
        {
            if (k == null) return false;
            return k.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string FindFirst(List<KeyValuePair<string, string>> props, string key)
        {
            if (props == null) return "";
            foreach (var kv in props)
            {
                if (IsKey(kv.Key, key)) return kv.Value ?? "";
            }
            return "";
        }

        private static long TryToInt64(object v)
        {
            if (v == null) return 0;
            try
            {
                if (v is long l) return l;
                if (v is int i) return i;
                if (v is short s) return s;
                if (v is byte b) return b;
                if (v is decimal d) return (long)d;
                if (v is double dd) return (long)dd;
                if (v is float ff) return (long)ff;

                var s0 = v.ToString();
                if (long.TryParse(s0, out var x)) return x;
            }
            catch { }
            return 0;
        }

        private static string TryToString(object v)
        {
            try { return v == null ? "" : v.ToString(); }
            catch { return ""; }
        }
    }
}

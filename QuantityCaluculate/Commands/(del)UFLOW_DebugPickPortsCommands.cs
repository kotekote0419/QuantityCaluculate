// UFLOW_DebugPickPortsCommands_v5_SubParts.cs
// Debug command: UFLOW_DEBUG_PICK_PORTS
// v5:
// - Add Connector SubParts DLM property inspection (AllSubParts + MakeAcPpObjectId + FindAcPpRowId + GetAllProperties)
// - Print KeyProps for LineTag / PartFamilyLongDesc / Size / InstallType with found flags
// - Print KeySearch(contains) to handle key name variations
//
// v4:
// - Keep robust DLM property enumeration (v3)
// - Improve InstallLengthService discovery: scan all loaded assemblies for type named "InstallLengthService"
// - If TryGetPipeEnds is unavailable, compute ends from Ports (S1/S2) and print both.
// - Also prints a quick mm->m hint if LengthUnit suggests "mm".

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Reflection;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFLOW
{
    public class UFLOW_DebugPickPortsCommands
    {
        [CommandMethod("UFLOW_DEBUG_PICK_PORTS")]
        public void DebugPickPorts()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick entity (Pipe/Part/Connector/PipeInlineAsset): ");
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
                        Log(ed, "[UFLOW][DBG] Selected object is not an Entity.");
                        return;
                    }

                    Log(ed, $"[UFLOW][DBG] Picked: Handle={ent.Handle} Type={ent.GetType().FullName}");

                    // ---- DLM props ----
                    DataLinksManager dlm = null;
                    var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    int rowId = -1;
                    string lengthUnit = "";

                    try
                    {
                        dlm = TryGetDataLinksManager();
                        if (dlm == null)
                        {
                            Log(ed, "[UFLOW][DBG] DataLinksManager not found via ProjectParts['Piping'].");
                        }
                        else
                        {
                            try { rowId = dlm.FindAcPpRowId(per.ObjectId); } catch { /* ignore */ }
                            props = SafeGetAllProperties(dlm, per.ObjectId, rowId);

                            string pnpClass = GetProp(props, "PnPClassName");
                            string pnpGuid = GetProp(props, "PnPGuid");
                            string lineTag = GetProp(props, "LineTag");
                            if (string.IsNullOrWhiteSpace(lineTag)) lineTag = GetProp(props, "LineNumberTag");
                            string spec = GetProp(props, "Spec");
                            string size = GetProp(props, "Size");
                            lengthUnit = GetProp(props, "LengthUnit");

                            Log(ed, $"[UFLOW][DBG] RowId={rowId} PnPClassName='{pnpClass}' PnPGuid='{pnpGuid}' LineTag='{lineTag}' Spec='{spec}' Size='{size}' LengthUnit='{lengthUnit}'");

                            Log(ed, "[UFLOW][DBG] ---- DLM property sample (first 40) ----");
                            foreach (var kv in props.Take(40))
                                Log(ed, $"[UFLOW][DBG]  {kv.Key} = {kv.Value}");

                            // ---- Connector SubParts (P3dConnector) ----
                            try
                            {
                                string pnpClassNameForSub = GetProp(props, "PnPClassName");
                                if (dlm != null && !string.IsNullOrWhiteSpace(pnpClassNameForSub) &&
                                    string.Equals(pnpClassNameForSub.Trim(), "P3dConnector", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Try to locate related "Accessory/Gasket" (or other) rows by walking OwnerId chain and dumping candidate properties.

                                    // Explore PnPDataLinks bundle (same DwgHandleLow/High) to see if Accessory/Gasket rows exist for this connector.
                                    try
                                    {
                                        DumpPnPDataLinksBundle(ed, dlm, rowId);
                                    }
                                    catch { /* ignore */ }

                                    try
                                    {
                                        DumpOwnerChainCandidates(ed, dlm, per.ObjectId, tr);
                                    }
                                    catch { /* ignore */ }

                                    LogConnectorSubParts(ed, dlm, per.ObjectId, ent);
                                }
                            }
                            catch { /* keep debug command robust */ }


                        }
                    }
                    catch (System.Exception ex)
                    {
                        Log(ed, "[UFLOW][DBG] DLM read failed: " + ex.Message);
                    }

                    // ---- Ports ----
                    var ports = TryGetPorts(ent, out string portsHow);
                    dynamic s1 = null;
                    dynamic s2 = null;

                    if (ports == null)
                    {
                        Log(ed, $"[UFLOW][DBG] Ports: (null) how={portsHow}");
                    }
                    else
                    {
                        Log(ed, $"[UFLOW][DBG] Ports count={ports.Count} how={portsHow}");
                        for (int i = 0; i < ports.Count; i++)
                        {
                            var portObj = ports[i];
                            string name = SafeGetString(portObj, "Name");
                            var pos = SafeGetPoint3d(portObj, "Position");

                            if (i == 0) s1 = pos;
                            if (i == 1) s2 = pos;

                            Log(ed, $"[UFLOW][DBG]  Port[{i}] Name='{name}' Pos=({Fmt(pos.X)},{Fmt(pos.Y)},{Fmt(pos.Z)})");
                        }

                        if (ports.Count >= 2)
                        {
                            var p0 = SafeGetPoint3d(ports[0], "Position");
                            var p1 = SafeGetPoint3d(ports[1], "Position");
                            double d = Dist(p0, p1);
                            Log(ed, $"[UFLOW][DBG] Distance(ports[0]-ports[1]) = {Fmt(d)} (drawing unit)");
                            if (IsMm(lengthUnit))
                                Log(ed, $"[UFLOW][DBG]  -> approx {Fmt(d / 1000.0)} m (assuming mm)");
                        }
                    }

                    // ---- TryGetPipeEnds (if available) ----
                    bool tryPipeEndsOk = false;
                    try
                    {
                        var ils = FindTypeByName("InstallLengthService");
                        if (ils != null)
                        {
                            var m = ils.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                                       .FirstOrDefault(x => x.Name == "TryGetPipeEnds" && x.GetParameters().Length == 3);
                            if (m != null)
                            {
                                object[] args = new object[] { ent, null, null };
                                bool ok = (bool)m.Invoke(null, args);
                                if (ok)
                                {
                                    tryPipeEndsOk = true;
                                    var pa = ToPoint3d(args[1]);
                                    var pb = ToPoint3d(args[2]);
                                    double d = Dist(pa, pb);
                                    Log(ed, $"[UFLOW][DBG] TryGetPipeEnds({ils.FullName}): OK A=({Fmt(pa.X)},{Fmt(pa.Y)},{Fmt(pa.Z)}) B=({Fmt(pb.X)},{Fmt(pb.Y)},{Fmt(pb.Z)}) dist={Fmt(d)}");
                                    if (IsMm(lengthUnit))
                                        Log(ed, $"[UFLOW][DBG]  -> approx {Fmt(d / 1000.0)} m (assuming mm)");
                                }
                                else
                                {
                                    Log(ed, $"[UFLOW][DBG] TryGetPipeEnds({ils.FullName}): FALSE");
                                }
                            }
                            else
                            {
                                Log(ed, $"[UFLOW][DBG] TryGetPipeEnds not found on {ils.FullName}.");
                            }
                        }
                        else
                        {
                            Log(ed, "[UFLOW][DBG] InstallLengthService type not found (scan all assemblies).");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Log(ed, "[UFLOW][DBG] TryGetPipeEnds invoke failed: " + ex.Message);
                    }

                    // ---- If TryGetPipeEnds unavailable, show a Ports-based "PipeEnds" candidate ----
                    if (!tryPipeEndsOk && s1 != null && s2 != null)
                    {
                        double d = Dist(s1, s2);
                        Log(ed, $"[UFLOW][DBG] PipeEnds candidate from Ports: A=({Fmt(s1.X)},{Fmt(s1.Y)},{Fmt(s1.Z)}) B=({Fmt(s2.X)},{Fmt(s2.Y)},{Fmt(s2.Z)}) dist={Fmt(d)}");
                        if (IsMm(lengthUnit))
                            Log(ed, $"[UFLOW][DBG]  -> approx {Fmt(d / 1000.0)} m (assuming mm)");
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                var d = Application.DocumentManager.MdiActiveDocument;
                d?.Editor?.WriteMessage("\n[UFLOW][DBG] ERROR: " + ex);
            }
        }

        private static DataLinksManager TryGetDataLinksManager()
        {
            try
            {
                var proj = PlantApplication.CurrentProject;
                if (proj == null) return null;

                object piping = null;
                try
                {
                    var pp = proj.ProjectParts;
                    if (pp != null)
                    {
                        var idx = pp.GetType().GetProperty("Item", new Type[] { typeof(string) });
                        if (idx != null) piping = idx.GetValue(pp, new object[] { "Piping" });
                    }
                }
                catch { }

                if (piping == null) return null;

                var pDlm = piping.GetType().GetProperty("DataLinksManager", BindingFlags.Public | BindingFlags.Instance);
                if (pDlm == null) return null;

                return pDlm.GetValue(piping, null) as DataLinksManager;
            }
            catch
            {
                return null;
            }
        }



        private static void DumpOwnerChainCandidates(Editor ed, DataLinksManager dlm, ObjectId startOid, Transaction tr)
        {
            const int MAX_DEPTH = 12;

            Log(ed, "[UFLOW][DBG] ---- OwnerId chain candidate rows ----");

            ObjectId cur = startOid;
            for (int depth = 0; depth < MAX_DEPTH; depth++)
            {
                if (cur == ObjectId.Null) break;

                int rid = -1;
                try { rid = dlm.FindAcPpRowId(cur); } catch { rid = -1; }

                Log(ed, $"[UFLOW][DBG] [Owner depth={depth}] oid={cur} rowId={rid}");

                if (rid > 0)
                {
                    TryRowIdAndDump(ed, dlm, cur, rid, $"Owner depth={depth}", maxPropsDump: 120);
                }

                // Walk to OwnerId
                try
                {
                    DBObject dbo = tr.GetObject(cur, OpenMode.ForRead, false);
                    if (dbo == null) break;

                    ObjectId next = ObjectId.Null;
                    try { next = dbo.OwnerId; } catch { next = ObjectId.Null; }

                    if (next == ObjectId.Null || next == cur) break;
                    cur = next;
                }
                catch
                {
                    break;
                }
            }

            Log(ed, "[UFLOW][DBG] ---- OwnerId chain end ----");
        }

        private static void TryRowIdAndDump(Editor ed, DataLinksManager dlm, ObjectId oid, int rowId, string label, int maxPropsDump = 120)
        {
            try
            {
                var props = SafeGetAllProperties(dlm, oid, rowId);
                if (props == null || props.Count == 0)
                {
                    Log(ed, $"[UFLOW][DBG] [{label}] rowId={rowId} props=0");
                    return;
                }

                string pnpClass = GetFirst(props, "PnPClassName");
                string lineTag = GetFirst(props, "LineTag", "LineNumber", "LineNo", "LINETAG");
                string size = GetFirst(props, "Size", "NominalSize", "PartSize");
                string installType = GetFirst(props, "InstallType", "InstallationType");
                string desc = GetFirst(props, "PartFamilyLongDesc", "PartFamilyLongDescription", "LongDescription", "Description");

                Log(ed, $"[UFLOW][DBG] [{label}] rowId={rowId} PnPClassName='{pnpClass}' LineTag='{lineTag}' Size='{size}' InstallType='{installType}' Desc='{desc}'");

                Log(ed, $"[UFLOW][DBG] [{label}] ---- DLM property dump (first {maxPropsDump}) ----");
                int n = 0;
                foreach (var kv in props.OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase))
                {
                    Log(ed, $"[UFLOW][DBG] [{label}]  {kv.Key} = {kv.Value}");
                    n++;
                    if (n >= maxPropsDump) break;
                }
            }
            catch (System.Exception ex)
            {
                Log(ed, $"[UFLOW][DBG] [{label}] dump failed: {ex.Message}");
            }
        }

        private static string GetFirst(Dictionary<string, string> props, params string[] keys)
        {
            if (props == null || keys == null) return "";
            foreach (var k in keys)
            {
                if (k == null) continue;
                if (props.TryGetValue(k, out var v) && v != null) return v;
            }
            return "";
        }


        // ---- PnPDataLinks bundle exploration (best-effort, reflection based) ----

        private static object TryGetPipingProjectPart()
        {
            try
            {
                var proj = PlantApplication.CurrentProject;
                if (proj == null) return null;

                var pp = proj.ProjectParts;
                if (pp == null) return null;

                var idx = pp.GetType().GetProperty("Item", new Type[] { typeof(string) });
                if (idx == null) return null;

                return idx.GetValue(pp, new object[] { "Piping" });
            }
            catch { return null; }
        }
        private static object TryGetPnPDatabase(object pipingProjectPart, DataLinksManager dlm, Editor ed)
        {
            // Plant 3D 2026 may keep PnPDatabase behind non-public members.
            // We scan (public+nonpublic) properties/fields/methods and accept ONLY real PnPDatabase (not *Mode).
            try
            {
                var flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;

                object Scan(object obj, string label)
                {
                    if (obj == null) return null;

                    var t = obj.GetType();
                    Log(ed, $"[UFLOW][DBG] {label} type={t.FullName}");

                    // properties
                    foreach (var p in t.GetProperties(flags))
                    {
                        if (p == null || !p.CanRead) continue;

                        object v = null;
                        try { v = p.GetValue(obj, null); } catch { continue; }
                        if (v == null) continue;

                        var vn = v.GetType().FullName ?? "";
                        if (vn.IndexOf("PnPDatabase", StringComparison.OrdinalIgnoreCase) < 0) continue;

                        Log(ed, $"[UFLOW][DBG] Candidate via {label}.{p.Name} type={vn}");
                        if (IsRealPnPDatabase(v))
                        {
                            Log(ed, "[UFLOW][DBG]   -> accepted as PnPDatabase");
                            return v;
                        }
                        Log(ed, "[UFLOW][DBG]   -> rejected (not real PnPDatabase)");
                    }

                    // fields
                    foreach (var f in t.GetFields(flags))
                    {
                        if (f == null) continue;

                        object v = null;
                        try { v = f.GetValue(obj); } catch { continue; }
                        if (v == null) continue;

                        var vn = v.GetType().FullName ?? "";
                        if (vn.IndexOf("PnPDatabase", StringComparison.OrdinalIgnoreCase) < 0) continue;

                        Log(ed, $"[UFLOW][DBG] Candidate via {label}.{f.Name} (field) type={vn}");
                        if (IsRealPnPDatabase(v))
                        {
                            Log(ed, "[UFLOW][DBG]   -> accepted as PnPDatabase");
                            return v;
                        }
                        Log(ed, "[UFLOW][DBG]   -> rejected (not real PnPDatabase)");
                    }

                    // methods (parameterless) returning PnPDatabase*
                    foreach (var mi in t.GetMethods(flags))
                    {
                        if (mi == null) continue;
                        if (mi.GetParameters().Length != 0) continue;

                        var rt = mi.ReturnType;
                        var rn = rt?.FullName ?? "";
                        if (rn.IndexOf("PnPDatabase", StringComparison.OrdinalIgnoreCase) < 0) continue;

                        object v = null;
                        try { v = mi.Invoke(obj, null); } catch { continue; }
                        if (v == null) continue;

                        var vn = v.GetType().FullName ?? "";
                        Log(ed, $"[UFLOW][DBG] Candidate via {label}.{mi.Name}() type={vn}");
                        if (IsRealPnPDatabase(v))
                        {
                            Log(ed, "[UFLOW][DBG]   -> accepted as PnPDatabase");
                            return v;
                        }
                        Log(ed, "[UFLOW][DBG]   -> rejected (not real PnPDatabase)");
                    }

                    return null;
                }

                var db1 = Scan(pipingProjectPart, "PipingProjectPart");
                if (db1 != null) return db1;

                var db2 = Scan(dlm, "DataLinksManager");
                if (db2 != null) return db2;

                return null;
            }
            catch
            {
                return null;
            }
        }
        private static bool IsRealPnPDatabase(object candidate)
        {
            if (candidate == null) return false;

            try
            {
                var t = candidate.GetType();
                var tn = (t.FullName ?? t.Name ?? "");

                // Reject obvious false positives (Mode objects)
                if (tn.IndexOf("PnPDatabaseMode", StringComparison.OrdinalIgnoreCase) >= 0) return false;
                if (tn.EndsWith("Mode", StringComparison.OrdinalIgnoreCase)) return false;

                var flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;

                // Must have Tables (any visibility) OR GetRowTableName(int)
                var pTables = t.GetProperty("Tables", flags);
                if (pTables != null)
                {
                    try
                    {
                        var tables = pTables.GetValue(candidate, null);
                        if (tables != null) return true;
                    }
                    catch { }
                }

                var m = t.GetMethod("GetRowTableName", flags, null, new[] { typeof(int) }, null);
                if (m != null) return true;

                return false;
            }
            catch
            {
                return false;
            }
        }
        private static void DumpPnPDatabaseTableNames(Editor ed, object pnpDb, int maxNames = 80)
        {
            try
            {
                if (pnpDb == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDatabase table dump: pnpDb is null");
                    return;
                }

                var flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;

                var t = pnpDb.GetType();
                var pTables = t.GetProperty("Tables", flags);
                if (pTables == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDatabase table dump: no Tables property");
                    return;
                }

                var tables = pTables.GetValue(pnpDb, null);
                if (tables == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDatabase table dump: Tables is null");
                    return;
                }

                Log(ed, $"[UFLOW][DBG] PnPDatabase.Tables type={tables.GetType().FullName}");

                // Prefer Keys if present
                var pKeys = tables.GetType().GetProperty("Keys", flags);
                var keys = pKeys?.GetValue(tables, null) as System.Collections.IEnumerable;
                if (keys != null)
                {
                    int n = 0;
                    foreach (var k in keys)
                    {
                        Log(ed, $"[UFLOW][DBG] TableKey[{++n}] = {k}");
                        if (n >= maxNames) break;
                    }
                    return;
                }

                // Otherwise enumerate
                if (tables is System.Collections.IEnumerable en)
                {
                    int n = 0;
                    foreach (var item in en)
                    {
                        if (item == null) continue;
                        string name = "";

                        try
                        {
                            var pn = item.GetType().GetProperty("Name", flags);
                            if (pn != null) name = (pn.GetValue(item, null)?.ToString() ?? "");
                        }
                        catch { }

                        if (string.IsNullOrWhiteSpace(name))
                            name = item.ToString() ?? "";

                        Log(ed, $"[UFLOW][DBG] Table[{++n}] = {name}");
                        if (n >= maxNames) break;
                    }
                    return;
                }

                Log(ed, "[UFLOW][DBG] PnPDatabase table dump: cannot enumerate tables");
            }
            catch (System.Exception ex)
            {
                Log(ed, "[UFLOW][DBG] PnPDatabase table dump failed: " + ex.Message);
            }
        }
        private static object TryGetPnPTable(object pnpDb, string tableName, Editor ed = null)
        {
            try
            {
                if (pnpDb == null || string.IsNullOrWhiteSpace(tableName)) return null;

                var flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;

                var pTables = pnpDb.GetType().GetProperty("Tables", flags);
                if (pTables == null) return null;

                var tables = pTables.GetValue(pnpDb, null);
                if (tables == null) return null;

                // Indexer: Tables["Name"]
                var idx = tables.GetType().GetProperty(
                    "Item",
                    flags,
                    null,
                    null,
                    new[] { typeof(string) },
                    null
                );

                // 1) exact name
                if (idx != null)
                {
                    try
                    {
                        var t = idx.GetValue(tables, new object[] { tableName });
                        if (t != null) return t;
                    }
                    catch { }
                }

                // 2) keys (case-insensitive / contains)
                var keysProp = tables.GetType().GetProperty("Keys", flags);
                var keys = keysProp?.GetValue(tables, null) as System.Collections.IEnumerable;

                if (keys != null && idx != null)
                {
                    foreach (var k in keys)
                    {
                        var ks = k?.ToString() ?? "";
                        if (ks.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            try { return idx.GetValue(tables, new object[] { ks }); } catch { }
                        }
                    }

                    foreach (var k in keys)
                    {
                        var ks = k?.ToString() ?? "";
                        bool contains = ks.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) >= 0;

                        bool datalinksAlias =
                            tableName.Equals("PnPDataLinks", StringComparison.OrdinalIgnoreCase) &&
                            (ks.IndexOf("DataLink", StringComparison.OrdinalIgnoreCase) >= 0 || ks.IndexOf("Link", StringComparison.OrdinalIgnoreCase) >= 0);

                        if (contains || datalinksAlias)
                        {
                            try
                            {
                                var t = idx.GetValue(tables, new object[] { ks });
                                if (t != null)
                                {
                                    if (ed != null) Log(ed, $"[UFLOW][DBG] TryGetPnPTable: using '{ks}' for requested '{tableName}'");
                                    return t;
                                }
                            }
                            catch { }
                        }
                    }
                }

                // 3) enumerate table objects with Name
                if (tables is System.Collections.IEnumerable en)
                {
                    foreach (var item in en)
                    {
                        if (item == null) continue;
                        string name = "";

                        try
                        {
                            var pn = item.GetType().GetProperty("Name", flags);
                            if (pn != null) name = (pn.GetValue(item, null)?.ToString() ?? "");
                        }
                        catch { }

                        if (string.IsNullOrWhiteSpace(name)) continue;

                        bool match = name.Equals(tableName, StringComparison.OrdinalIgnoreCase) ||
                                     name.IndexOf(tableName, StringComparison.OrdinalIgnoreCase) >= 0;

                        bool datalinksAlias =
                            tableName.Equals("PnPDataLinks", StringComparison.OrdinalIgnoreCase) &&
                            (name.IndexOf("DataLink", StringComparison.OrdinalIgnoreCase) >= 0 || name.IndexOf("Link", StringComparison.OrdinalIgnoreCase) >= 0);

                        if (match || datalinksAlias) return item;
                    }
                }

                if (ed != null)
                {
                    Log(ed, $"[UFLOW][DBG] TryGetPnPTable: '{tableName}' not found. Dumping table names for inspection...");
                    DumpPnPDatabaseTableNames(ed, pnpDb, 80);
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        private static System.Collections.IEnumerable TrySelectRows(object pnpTable, string where)
        {
            if (pnpTable == null) return null;

            try
            {
                // Common signatures: Select(string), Select(string, string), Select(string, bool)
                var t = pnpTable.GetType();

                var m1 = t.GetMethod("Select", new Type[] { typeof(string) });
                if (m1 != null) return m1.Invoke(pnpTable, new object[] { where }) as System.Collections.IEnumerable;

                var m2 = t.GetMethod("Select", new Type[] { typeof(string), typeof(string) });
                if (m2 != null) return m2.Invoke(pnpTable, new object[] { where, "" }) as System.Collections.IEnumerable;

                var m3 = t.GetMethod("Select", new Type[] { typeof(string), typeof(bool) });
                if (m3 != null) return m3.Invoke(pnpTable, new object[] { where, true }) as System.Collections.IEnumerable;
            }
            catch { }

            return null;
        }


        private static List<string> TryGetPnPTableColumnNames(object pnpTable)
        {
            var names = new List<string>();
            if (pnpTable == null) return names;

            try
            {
                var flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
                var t = pnpTable.GetType();

                // Common: pnpTable.Columns
                var pCols = t.GetProperty("Columns", flags);
                if (pCols != null)
                {
                    var cols = pCols.GetValue(pnpTable, null) as System.Collections.IEnumerable;
                    if (cols != null)
                    {
                        foreach (var c in cols)
                        {
                            if (c == null) continue;
                            string name = "";
                            try
                            {
                                var pn = c.GetType().GetProperty("Name", flags);
                                if (pn != null) name = (pn.GetValue(c, null)?.ToString() ?? "");
                            }
                            catch { }
                            if (string.IsNullOrWhiteSpace(name)) name = c.ToString() ?? "";
                            if (!string.IsNullOrWhiteSpace(name)) names.Add(name);
                        }
                    }
                }

                // Fallback: pnpTable.ColumnNames
                if (names.Count == 0)
                {
                    var p = t.GetProperty("ColumnNames", flags);
                    if (p != null)
                    {
                        var en = p.GetValue(pnpTable, null) as System.Collections.IEnumerable;
                        if (en != null)
                        {
                            foreach (var x in en)
                            {
                                var s = x?.ToString() ?? "";
                                if (!string.IsNullOrWhiteSpace(s)) names.Add(s);
                            }
                        }
                    }
                }

                // De-dupe
                names = names.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            }
            catch { }

            return names;
        }

        private static void DumpPnPRowAllColumns(Editor ed, object pnpTable, object row, string prefix, int maxCols = 200)
        {
            if (ed == null || row == null) return;

            try
            {
                var cols = TryGetPnPTableColumnNames(pnpTable);
                if (cols.Count == 0)
                {
                    Log(ed, $"[UFLOW][DBG] {prefix} columns: (cannot enumerate columns from table; dumping public properties instead)");
                    var flags = BindingFlags.Public | BindingFlags.Instance;
                    int n = 0;
                    foreach (var p in row.GetType().GetProperties(flags))
                    {
                        if (!p.CanRead) continue;
                        object v = null;
                        try { v = p.GetValue(row, null); } catch { continue; }
                        Log(ed, $"[UFLOW][DBG] {prefix}  {p.Name} = {v}");
                        if (++n >= maxCols) break;
                    }
                    return;
                }

                Log(ed, $"[UFLOW][DBG] {prefix} ---- PnPRow column dump (first {Math.Min(maxCols, cols.Count)}) ----");
                int c = 0;
                foreach (var name in cols)
                {
                    object v = null;
                    try { v = TryGetRowField(row, name); } catch { v = null; }
                    Log(ed, $"[UFLOW][DBG] {prefix}  {name} = {v}");
                    if (++c >= maxCols) break;
                }
            }
            catch (System.Exception ex)
            {
                Log(ed, $"[UFLOW][DBG] {prefix} column dump failed: {ex.Message}");
            }
        }

        private static List<int> ExtractRowIdCandidatesFromPnPRow(object pnpTable, object row, int maxScan = 220)
        {
            var ids = new List<int>();
            if (row == null) return ids;

            try
            {
                var cols = TryGetPnPTableColumnNames(pnpTable);
                if (cols.Count == 0) return ids;

                // prioritize likely columns first
                var ordered = cols
                    .OrderByDescending(n =>
                    {
                        var s = n.ToLowerInvariant();
                        int score = 0;
                        if (s.Contains("rowid")) score += 5;
                        if (s.Contains("pnp")) score += 3;
                        if (s.Contains("ref") || s.Contains("related") || s.Contains("target") || s.Contains("source")) score += 2;
                        if (s.Contains("id")) score += 1;
                        return score;
                    })
                    .ToList();

                int scanned = 0;
                foreach (var name in ordered)
                {
                    object v = null;
                    try { v = TryGetRowField(row, name); } catch { v = null; }
                    scanned++;
                    if (v == null) continue;

                    long x = TryToInt64(v);
                    if (x > 0 && x < int.MaxValue)
                    {
                        // heuristics: ignore dwghandle low/high and timestamps
                        var s = name.ToLowerInvariant();
                        if (s.Contains("dwghandle")) continue;
                        if (s.Contains("timestamp")) continue;

                        ids.Add((int)x);
                    }

                    if (scanned >= maxScan) break;
                }

                ids = ids.Distinct().ToList();
            }
            catch { }

            return ids;
        }

        private static void TryResolveCandidateRowIds(Editor ed, object pnpDb, DataLinksManager dlm, IEnumerable<int> candidateIds, string prefix, int limit = 30)
        {
            if (ed == null || candidateIds == null) return;

            int n = 0;
            foreach (var rid in candidateIds)
            {
                if (rid <= 0) continue;
                if (++n > limit) break;

                string tableName = "";
                try
                {
                    var flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
                    var m = pnpDb?.GetType().GetMethod("GetRowTableName", flags, null, new Type[] { typeof(int) }, null);
                    if (m != null)
                    {
                        tableName = m.Invoke(pnpDb, new object[] { rid })?.ToString() ?? "";
                    }
                }
                catch { }

                // Also attempt DLM props for this rowId (if accessible)
                string cls = "", lt = "", sz = "", it = "";
                try
                {
                    var props = SafeGetAllProperties(dlm, ObjectId.Null, rid);
                    if (props != null)
                    {
                        props.TryGetValue("PnPClassName", out cls);
                        props.TryGetValue("LineTag", out lt);
                        props.TryGetValue("Size", out sz);
                        props.TryGetValue("InstallType", out it);
                    }
                }
                catch { }

                Log(ed, $"[UFLOW][DBG] {prefix} candRowId={rid} Table='{tableName}' PnPClassName='{cls}' LineTag='{lt}' Size='{sz}' InstallType='{it}'");
            }
        }

        private static object TryGetRowField(object row, string fieldName)
        {
            if (row == null || string.IsNullOrWhiteSpace(fieldName)) return null;

            try
            {
                // Indexer: row["ColName"]
                var idx = row.GetType().GetProperty("Item", new Type[] { typeof(string) });
                if (idx != null)
                {
                    try { return idx.GetValue(row, new object[] { fieldName }); } catch { }
                }

                // Property: ColName
                var p = row.GetType().GetProperty(fieldName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (p != null && p.CanRead)
                {
                    try { return p.GetValue(row, null); } catch { }
                }
            }
            catch { }

            return null;
        }

        private static long TryToInt64(object v)
        {
            try
            {
                if (v == null) return 0;
                if (v is long l) return l;
                if (v is int i) return i;
                if (v is short s) return s;
                if (v is string str && long.TryParse(str, out var x)) return x;
                return System.Convert.ToInt64(v);
            }
            catch { return 0; }
        }


        private class PnPDataLinkRowInfo
        {
            public long LinkPnPID;
            public int RowId;
            public int DwgId;
            public int DwgSubIndex;
            public string RowClassName;
        }

        private static int TryToInt32(object v)
        {
            try
            {
                if (v == null) return 0;
                if (v is int i) return i;
                if (v is long l && l <= int.MaxValue) return (int)l;
                if (v is short s) return s;
                if (v is string str && int.TryParse(str, out var x)) return x;
                return System.Convert.ToInt32(v);
            }
            catch { return 0; }
        }

        private static string TryToString(object v)
        {
            try { return v?.ToString() ?? ""; } catch { return ""; }
        }

        private static void DumpKeyProps(Editor ed, DataLinksManager dlm, int rowId, string prefix, int maxProps = 160)
        {
            if (ed == null || dlm == null || rowId <= 0) return;

            try
            {
                var props = SafeGetAllProperties(dlm, ObjectId.Null, rowId);
                if (props == null || props.Count == 0)
                {
                    Log(ed, $"[UFLOW][DBG] {prefix} DLM props: (none) rowId={rowId}");
                    return;
                }

                // Always print a few important keys if present
                string[] keys = new[]
                {
                    "PnPClassName","PnPGuid","LineTag","Size","Spec","ShortDesc","LongDesc","PartFamilyLongDesc",
                    "InstallType","JointType","Tag","PartFamily","PartFamilyDesc"
                };

                foreach (var k in keys)
                {
                    if (props.TryGetValue(k, out var val))
                        Log(ed, $"[UFLOW][DBG] {prefix} {k} = {val}");
                }

                // Then dump the first N properties (stable order)
                int n = 0;
                Log(ed, $"[UFLOW][DBG] {prefix} ---- DLM property dump (first {Math.Min(maxProps, props.Count)}) ----");
                foreach (var kv in props)
                {
                    Log(ed, $"[UFLOW][DBG] {prefix}  {kv.Key} = {kv.Value}");
                    if (++n >= maxProps) break;
                }
            }
            catch (System.Exception ex)
            {
                Log(ed, $"[UFLOW][DBG] {prefix} DLM prop dump failed: {ex.Message}");
            }
        }

        private static void DumpPnPDataLinksBundle(Editor ed, DataLinksManager dlm, int baseRowId)
        {
            try
            {
                Log(ed, "[UFLOW][DBG] ---- PnPDataLinks bundle (same DwgHandleLow/High) ----");

                var piping = TryGetPipingProjectPart();
                var db = TryGetPnPDatabase(piping, dlm, ed);

                if (db == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDatabase not available (no public access). Dumping DB-like members to help adjust reflection...");

                    try
                    {
                        if (piping != null)
                        {
                            var tp = piping.GetType();
                            foreach (var p in tp.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                            {
                                if (p == null) continue;
                                var n = p.Name ?? "";
                                if (n.IndexOf("db", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("database", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("pnp", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    Log(ed, $"[UFLOW][DBG] piping prop: {p.Name} : {(p.PropertyType?.FullName ?? "")}");
                                }
                            }

                            foreach (var mi in tp.GetMethods(BindingFlags.Public | BindingFlags.Instance))
                            {
                                if (mi == null) continue;
                                if (mi.GetParameters().Length != 0) continue;
                                var n = mi.Name ?? "";
                                if (n.IndexOf("db", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("database", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("pnp", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    Log(ed, $"[UFLOW][DBG] piping method: {mi.Name}() -> {(mi.ReturnType?.FullName ?? "")}");
                                }
                            }
                        }

                        if (dlm != null)
                        {
                            var td = dlm.GetType();
                            foreach (var p in td.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                            {
                                if (p == null) continue;
                                var n = p.Name ?? "";
                                if (n.IndexOf("db", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("database", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("pnp", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    Log(ed, $"[UFLOW][DBG] dlm prop: {p.Name} : {(p.PropertyType?.FullName ?? "")}");
                                }
                            }

                            foreach (var mi in td.GetMethods(BindingFlags.Public | BindingFlags.Instance))
                            {
                                if (mi == null) continue;
                                if (mi.GetParameters().Length != 0) continue;
                                var n = mi.Name ?? "";
                                if (n.IndexOf("db", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("database", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    n.IndexOf("pnp", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    Log(ed, $"[UFLOW][DBG] dlm method: {mi.Name}() -> {(mi.ReturnType?.FullName ?? "")}");
                                }
                            }
                        }
                    }
                    catch { }

                    return;
                }


                var links = TryGetPnPTable(db, "PnPDataLinks", ed);
                if (links == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDataLinks table not found in PnPDatabase");
                    return;
                }

                // Step 1: find DwgHandleLow/High rows for this baseRowId
                var whereCandidates = new[]
                {
                    $"PnPID={baseRowId}",
                    $"PnPRowId={baseRowId}",
                    $"RowId={baseRowId}",
                    $"PnPId={baseRowId}"
                };

                System.Collections.IEnumerable baseRows = null;
                string usedWhere = null;

                foreach (var w in whereCandidates)
                {
                    var rr = TrySelectRows(links, w);
                    if (rr != null)
                    {
                        // test if any rows
                        var en = rr.GetEnumerator();
                        if (en != null && en.MoveNext())
                        {
                            baseRows = rr;
                            usedWhere = w;
                            break;
                        }
                    }
                }

                if (baseRows == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDataLinks: no rows found for this RowId by common keys (PnPID/PnPRowId/RowId).");
                    return;
                }

                Log(ed, $"[UFLOW][DBG] PnPDataLinks: baseRows found by where='{usedWhere}'");

                // pick the first baseRow to get handle keys (often identical across all baseRows)
                object first = null;
                foreach (var r in baseRows) { first = r; break; }
                if (first == null)
                {
                    Log(ed, "[UFLOW][DBG] PnPDataLinks: baseRows empty");
                    return;
                }

                var lowV = TryGetRowField(first, "DwgHandleLow");
                var highV = TryGetRowField(first, "DwgHandleHigh");

                long low = TryToInt64(lowV);
                long high = TryToInt64(highV);

                Log(ed, $"[UFLOW][DBG] PnPDataLinks: DwgHandleLow={low} DwgHandleHigh={high}");

                if (low == 0 && high == 0)
                {
                    Log(ed, "[UFLOW][DBG] PnPDataLinks: handle fields not found or zero; dump baseRow fields to inspect columns.");
                    // dump a few fields that might exist
                    var suspect = new[] { "DwgHandleLow", "DwgHandleHigh", "DwgSubIndex", "PnPID", "PnPRowId", "PnPGuid" };
                    foreach (var k in suspect)
                    {
                        var vv = TryGetRowField(first, k);
                        Log(ed, $"[UFLOW][DBG]  baseRow[{k}] = {vv}");
                    }
                    return;
                }

                // Step 2: get all rows with same handle
                string whereHandle = (high != 0)
                    ? $"DwgHandleLow={low} AND DwgHandleHigh={high}"
                    : $"DwgHandleLow={low}";

                var bundleRows = TrySelectRows(links, whereHandle);
                if (bundleRows == null)
                {
                    Log(ed, $"[UFLOW][DBG] PnPDataLinks: Select(whereHandle) returned null. where='{whereHandle}'");
                    return;
                }

                // Collect candidate RowIds from bundle
                var candidateRowIds = new System.Collections.Generic.List<int>();
                int rowCount = 0;


                var linkInfos = new List<PnPDataLinkRowInfo>();
                foreach (var r in bundleRows)
                {
                    rowCount++;

                    long pnpid = TryToInt64(TryGetRowField(r, "PnPID"));
                    if (pnpid == 0) pnpid = TryToInt64(TryGetRowField(r, "PnPRowId"));
                    if (pnpid == 0) pnpid = TryToInt64(TryGetRowField(r, "RowId"));

                    long sub = TryToInt64(TryGetRowField(r, "DwgSubIndex"));


                    // v14: capture link-row fields that point to real rows
                    int linkRowId = TryToInt32(TryGetRowField(r, "RowId"));
                    int linkDwgId = TryToInt32(TryGetRowField(r, "DwgId"));
                    string linkClass = TryToString(TryGetRowField(r, "RowClassName"));
                    int linkSub = TryToInt32(TryGetRowField(r, "DwgSubIndex"));

                    linkInfos.Add(new PnPDataLinkRowInfo
                    {
                        LinkPnPID = (long)pnpid,
                        RowId = linkRowId,
                        DwgId = linkDwgId,
                        DwgSubIndex = linkSub,
                        RowClassName = linkClass
                    });

                    Log(ed, $"[UFLOW][DBG] PnPDataLinks row[{rowCount}] PnPID={pnpid} DwgSubIndex={sub}");

                    // v13: dump link-row columns for first few rows, and try to resolve embedded RowId candidates
                    if (rowCount <= 6)
                    {
                        DumpPnPRowAllColumns(ed, links, r, $"[PnPDataLinks rowId? {pnpid}] ", maxCols: 220);
                    }

                    var embeddedIds = ExtractRowIdCandidatesFromPnPRow(links, r, maxScan: 240);
                    if (embeddedIds != null && embeddedIds.Count > 0)
                    {
                        // remove obvious self references and baseRowId
                        embeddedIds = embeddedIds.Where(x => x != baseRowId && x != (int)pnpid).Distinct().ToList();
                        if (embeddedIds.Count > 0)
                        {
                            Log(ed, $"[UFLOW][DBG] PnPDataLinks row[{rowCount}] embedded candidate ids count={embeddedIds.Count}");
                            TryResolveCandidateRowIds(ed, db, dlm, embeddedIds, $"[PnPDataLinks row[{rowCount}]]", limit: 25);
                        }
                    }


                    if (pnpid > 0 && pnpid < int.MaxValue) candidateRowIds.Add((int)pnpid);
                }

                candidateRowIds = candidateRowIds.Distinct().ToList();
                Log(ed, $"[UFLOW][DBG] PnPDataLinks: bundle candidate RowIds count={candidateRowIds.Count}");

                // v14: Determine baseDwgId from the link row that represents the picked baseRowId (usually RowClassName=P3dConnector, DwgSubIndex=0)
                int baseDwgId = 0;
                try
                {
                    var baseInfo = linkInfos.FirstOrDefault(x =>
                        x != null &&
                        x.RowId == baseRowId &&
                        (x.RowClassName ?? "").Equals("P3dConnector", StringComparison.OrdinalIgnoreCase) &&
                        x.DwgSubIndex == 0);

                    if (baseInfo == null)
                    {
                        // fallback: any row with RowId==baseRowId
                        baseInfo = linkInfos.FirstOrDefault(x => x != null && x.RowId == baseRowId);
                    }

                    baseDwgId = baseInfo != null ? baseInfo.DwgId : 0;

                    Log(ed, $"[UFLOW][DBG] v14 baseDwgId resolved = {baseDwgId} (baseRowId={baseRowId})");
                }
                catch { }

                // v14: Filter to same DwgId as base, and locate Gasket/BoltSet real rowIds
                try
                {
                    if (baseDwgId != 0)
                    {
                        var sameDwg = linkInfos.Where(x => x != null && x.DwgId == baseDwgId).ToList();
                        Log(ed, $"[UFLOW][DBG] v14 sameDwg link rows count={sameDwg.Count}");

                        var gaskets = sameDwg.Where(x => (x.RowClassName ?? "").Equals("Gasket", StringComparison.OrdinalIgnoreCase)).ToList();
                        var boltsets = sameDwg.Where(x => (x.RowClassName ?? "").Equals("BoltSet", StringComparison.OrdinalIgnoreCase)).ToList();

                        Log(ed, $"[UFLOW][DBG] v14 gaskets count={gaskets.Count}, boltsets count={boltsets.Count}");

                        int i = 0;
                        foreach (var g in gaskets)
                        {
                            Log(ed, $"[UFLOW][DBG] v14 Gasket[{++i}] realRowId={g.RowId} subIndex={g.DwgSubIndex}");
                            DumpKeyProps(ed, dlm, g.RowId, $"[Gasket realRowId={g.RowId}] ", maxProps: 200);
                        }

                        i = 0;
                        foreach (var b in boltsets)
                        {
                            Log(ed, $"[UFLOW][DBG] v14 BoltSet[{++i}] realRowId={b.RowId} subIndex={b.DwgSubIndex}");
                            DumpKeyProps(ed, dlm, b.RowId, $"[BoltSet realRowId={b.RowId}] ", maxProps: 200);
                        }
                    }
                    else
                    {
                        Log(ed, "[UFLOW][DBG] v14 baseDwgId could not be resolved; skipping gasket/boltset detailed dump.");
                    }
                }
                catch (System.Exception ex)
                {
                    Log(ed, "[UFLOW][DBG] v14 gasket/boltset dump failed: " + ex.Message);
                }


                // Step 3: For each candidate RowId, print table name (if possible) and key props
                foreach (var rid in candidateRowIds)
                {
                    string tableName = "";
                    try
                    {
                        var m = db.GetType().GetMethod("GetRowTableName", new Type[] { typeof(int) });
                        if (m != null)
                        {
                            var tn = m.Invoke(db, new object[] { rid });
                            tableName = tn?.ToString() ?? "";
                        }
                    }
                    catch { }

                    var props = SafeGetAllProperties(dlm, ObjectId.Null, rid); // SafeGetAllProperties ignores oid if rowId overload exists
                    string cls = props != null && props.TryGetValue("PnPClassName", out var c) ? c : "";
                    string lt = props != null && props.TryGetValue("LineTag", out var x) ? x : "";
                    string sz = props != null && props.TryGetValue("Size", out var y) ? y : "";
                    string it = props != null && props.TryGetValue("InstallType", out var z) ? z : "";

                    Log(ed, $"[UFLOW][DBG] BundleRowId={rid} Table='{tableName}' PnPClassName='{cls}' LineTag='{lt}' Size='{sz}' InstallType='{it}'");
                }

                Log(ed, "[UFLOW][DBG] ---- PnPDataLinks bundle end ----");
            }
            catch (System.Exception ex)
            {
                Log(ed, "[UFLOW][DBG] PnPDataLinks bundle failed: " + ex.Message);
            }
        }

        private static void DumpAllPropertiesToLog(Editor ed, object props, string prefix, int max = 500)
        {
            if (ed == null || props == null) return;
            int count = 0;

            if (props is System.Collections.IEnumerable en)
            {
                foreach (var item in en)
                {
                    if (item == null) continue;

                    string k = null;
                    string v = null;

                    var it = item.GetType();
                    var kpi = it.GetProperty("Key");
                    var vpi = it.GetProperty("Value");
                    if (kpi != null && vpi != null)
                    {
                        k = kpi.GetValue(item, null)?.ToString();
                        v = vpi.GetValue(item, null)?.ToString();
                    }
                    else
                    {
                        var keyProp = it.GetProperty("Name");
                        var valProp = it.GetProperty("Value");
                        if (keyProp != null && valProp != null)
                        {
                            k = keyProp.GetValue(item, null)?.ToString();
                            v = valProp.GetValue(item, null)?.ToString();
                        }
                    }

                    if (k != null)
                    {
                        Log(ed, $"[UFLOW][DBG] {prefix}{k} = {v}");
                        count++;
                        if (count >= max) break;
                    }
                }
            }
        }

        private static Dictionary<string, string> SafeGetAllProperties(DataLinksManager dlm, object acPpObj, int rowId)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            object p1 = null;
            try { p1 = TryGetAllPropertiesByObject(dlm, acPpObj); } catch { }
            MergeProps(dict, p1);

            if (rowId > 0)
            {
                object p2 = null;
                try { p2 = dlm.GetAllProperties(rowId, true); } catch { }
                MergeProps(dict, p2);
            }

            return dict;
        }

        private static void MergeProps(Dictionary<string, string> dict, object propsObj)
        {
            if (dict == null || propsObj == null) return;

            if (propsObj is NameValueCollection nvc)
            {
                foreach (string k in nvc.AllKeys)
                {
                    if (string.IsNullOrEmpty(k)) continue;
                    if (!dict.ContainsKey(k)) dict[k] = nvc[k] ?? "";
                }
                return;
            }

            if (propsObj is IDictionary id)
            {
                foreach (DictionaryEntry de in id)
                {
                    var k = Convert.ToString(de.Key, CultureInfo.InvariantCulture) ?? "";
                    if (string.IsNullOrEmpty(k)) continue;
                    if (!dict.ContainsKey(k)) dict[k] = Convert.ToString(de.Value, CultureInfo.InvariantCulture) ?? "";
                }
                return;
            }

            var t = propsObj.GetType();

            try
            {
                var pAllKeys = t.GetProperty("AllKeys", BindingFlags.Public | BindingFlags.Instance);
                if (pAllKeys != null)
                {
                    var keys = pAllKeys.GetValue(propsObj, null) as string[];
                    if (keys != null)
                    {
                        foreach (var k in keys)
                        {
                            if (string.IsNullOrEmpty(k)) continue;
                            if (!dict.ContainsKey(k))
                                dict[k] = GetValueByKey(propsObj, k);
                        }
                        return;
                    }
                }
            }
            catch { }

            try
            {
                var pKeys = t.GetProperty("Keys", BindingFlags.Public | BindingFlags.Instance);
                if (pKeys != null)
                {
                    var keysObj = pKeys.GetValue(propsObj, null);
                    if (keysObj is IEnumerable en)
                    {
                        foreach (var kk in en)
                        {
                            var k = Convert.ToString(kk, CultureInfo.InvariantCulture) ?? "";
                            if (string.IsNullOrEmpty(k)) continue;
                            if (!dict.ContainsKey(k))
                                dict[k] = GetValueByKey(propsObj, k);
                        }
                        return;
                    }
                }
            }
            catch { }

            try
            {
                if (propsObj is IEnumerable en2)
                {
                    foreach (var item in en2)
                    {
                        if (item == null) continue;
                        var it = item.GetType();
                        var pK = it.GetProperty("Key");
                        var pV = it.GetProperty("Value");
                        if (pK != null && pV != null)
                        {
                            var k = Convert.ToString(pK.GetValue(item, null), CultureInfo.InvariantCulture) ?? "";
                            if (string.IsNullOrEmpty(k)) continue;
                            if (!dict.ContainsKey(k))
                                dict[k] = Convert.ToString(pV.GetValue(item, null), CultureInfo.InvariantCulture) ?? "";
                        }
                    }
                }
            }
            catch { }
        }

        private static string GetValueByKey(object propsObj, string key)
        {
            if (propsObj == null || string.IsNullOrEmpty(key)) return "";

            try
            {
                var t = propsObj.GetType();
                var idx = t.GetProperty("Item", new Type[] { typeof(string) });
                if (idx != null)
                {
                    var v = idx.GetValue(propsObj, new object[] { key });
                    return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
                }

                var m = t.GetMethod("get_Item", new Type[] { typeof(string) });
                if (m != null)
                {
                    var v = m.Invoke(propsObj, new object[] { key });
                    return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
                }
            }
            catch { }

            return "";
        }

        private static string GetProp(Dictionary<string, string> props, string key)
            => props != null && props.TryGetValue(key, out var v) ? (v ?? "") : "";

        private static bool IsMm(string unit)
        {
            if (string.IsNullOrWhiteSpace(unit)) return false;
            unit = unit.Trim();
            return unit.Equals("mm", StringComparison.OrdinalIgnoreCase) ||
                   unit.Equals("millimeter", StringComparison.OrdinalIgnoreCase) ||
                   unit.Equals("millimetre", StringComparison.OrdinalIgnoreCase);
        }

        private static List<object> TryGetPorts(object ent, out string how)
        {
            how = "(none)";
            if (ent == null) return null;

            try
            {
                var t = ent.GetType();
                var m = t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                         .FirstOrDefault(x => x.Name == "GetPorts" && x.GetParameters().Length == 1);
                if (m != null)
                {
                    var pt = m.GetParameters()[0].ParameterType;
                    object portTypeStatic = null;
                    foreach (var v in Enum.GetValues(pt))
                    {
                        if (v.ToString().Equals("Static", StringComparison.OrdinalIgnoreCase))
                        {
                            portTypeStatic = v;
                            break;
                        }
                    }
                    if (portTypeStatic == null) portTypeStatic = Activator.CreateInstance(pt);

                    var portsObj = m.Invoke(ent, new object[] { portTypeStatic });
                    var list = ToList(portsObj);
                    if (list != null)
                    {
                        how = "GetPorts(PortType.Static)";
                        return list;
                    }
                }
            }
            catch { }

            try
            {
                var p = ent.GetType().GetProperty("Ports", BindingFlags.Public | BindingFlags.Instance);
                if (p != null)
                {
                    var portsObj = p.GetValue(ent, null);
                    var list = ToList(portsObj);
                    if (list != null)
                    {
                        how = "Property Ports";
                        return list;
                    }
                }
            }
            catch { }

            how = "not found";
            return null;
        }

        private static List<object> ToList(object enumerableObj)
        {
            if (enumerableObj == null) return null;
            if (enumerableObj is IEnumerable en)
            {
                var list = new List<object>();
                foreach (var x in en) list.Add(x);
                return list;
            }
            return null;
        }

        private static Type FindTypeByName(string typeName)
        {
            if (string.IsNullOrWhiteSpace(typeName)) return null;

            try
            {
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    Type t = null;
                    try
                    {
                        t = asm.GetTypes().FirstOrDefault(x =>
                            x.Name == typeName ||
                            (x.FullName != null && x.FullName.EndsWith("." + typeName, StringComparison.Ordinal)));
                    }
                    catch { }
                    if (t != null) return t;
                }
            }
            catch { }

            return null;
        }

        private static void Log(Editor ed, string msg) => ed.WriteMessage("\n" + msg);

        private static string Fmt(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

        private static double Dist(dynamic a, dynamic b)
        {
            double dx = (double)a.X - (double)b.X;
            double dy = (double)a.Y - (double)b.Y;
            double dz = (double)a.Z - (double)b.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        private static dynamic ToPoint3d(object obj)
        {
            if (obj == null) return new { X = 0.0, Y = 0.0, Z = 0.0 };
            double x = SafeGetDouble(obj, "X");
            double y = SafeGetDouble(obj, "Y");
            double z = SafeGetDouble(obj, "Z");
            return new { X = x, Y = y, Z = z };
        }

        private static string SafeGetString(object obj, string propName)
        {
            try
            {
                var p = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (p == null) return "";
                var v = p.GetValue(obj, null);
                return v != null ? v.ToString() : "";
            }
            catch { return ""; }
        }

        private static double SafeGetDouble(object obj, string propName)
        {
            try
            {
                var p = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (p == null) return 0.0;
                var v = p.GetValue(obj, null);
                if (v == null) return 0.0;
                if (v is double dd) return dd;
                if (double.TryParse(v.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return d;
                if (double.TryParse(v.ToString(), NumberStyles.Float, CultureInfo.CurrentCulture, out d)) return d;
                return 0.0;
            }
            catch { return 0.0; }
        }

        private static dynamic SafeGetPoint3d(object obj, string propName)
        {
            double x = 0, y = 0, z = 0;
            try
            {
                var p = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (p != null)
                {
                    var v = p.GetValue(obj, null);
                    if (v != null)
                    {
                        x = SafeGetDouble(v, "X");
                        y = SafeGetDouble(v, "Y");
                        z = SafeGetDouble(v, "Z");
                        return new { X = x, Y = y, Z = z };
                    }
                }
            }
            catch { }
            return new { X = x, Y = y, Z = z };
        }

        private static ObjectId SafeGetObjectId(object obj, string propName)
        {
            try
            {
                var p = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (p == null) return ObjectId.Null;
                var v = p.GetValue(obj, null);
                if (v is ObjectId oid) return oid;
            }
            catch { }
            return ObjectId.Null;
        }

        // =====================================================================
        // Connector SubParts logging
        // =====================================================================
        private static void LogConnectorSubParts(Editor ed, DataLinksManager dlm, ObjectId baseOid, Entity ent)
        {
            Log(ed, "[UFLOW][DBG] ---- Connector SubParts ----");

            object subPartsObj = null;
            try
            {
                if (ent is Connector c) subPartsObj = c.AllSubParts;
                else
                {
                    var pi = ent.GetType().GetProperty("AllSubParts", BindingFlags.Public | BindingFlags.Instance);
                    if (pi != null) subPartsObj = pi.GetValue(ent, null);
                }
            }
            catch { subPartsObj = null; }

            var subParts = new List<object>();
            try
            {
                if (subPartsObj is IEnumerable en)
                {
                    foreach (var x in en) subParts.Add(x);
                }
            }
            catch { /* ignore */ }

            Log(ed, $"[UFLOW][DBG] Connector.AllSubParts count = {subParts.Count}");

            if (subParts.Count == 0)
            {
                Log(ed, "[UFLOW][DBG] (No SubParts) -> cannot inspect sub-part DLM properties.");
                return;
            }

            bool hasAnyRowId = false;
            bool hasLineTag = false;
            bool hasDesc = false;
            bool hasSize = false;
            bool hasInstallType = false;

            for (int idx = 1; idx <= subParts.Count; idx++)
            {
                object subKeyObj = null;
                int rowId = -1;

                try
                {
                    subKeyObj = TryMakeAcPpObjectId(dlm, baseOid, idx);
                    if (subKeyObj != null)
                    {
                        rowId = TryFindRowIdFromAcPpObject(dlm, subKeyObj);
                    }
                }
                catch { /* ignore */ }

                bool ok = (rowId > 0);
                Log(ed, $"[UFLOW][DBG] [SubPart {idx}] acPpObject=MakeAcPpObjectId(baseOid,{idx}) type={(subKeyObj == null ? "<null>" : subKeyObj.GetType().FullName)} rowId={rowId} ok={ok}");

                if (!ok)
                {
                    // Extra diagnostics: sometimes SubPart DLM data is linked to the raw AllSubParts item, not to MakeAcPpObjectId(...)
                    try
                    {
                        var rawSub = subParts[idx - 1];
                        Log(ed, $"[UFLOW][DBG] [SubPart {idx}] rawSubPart type={(rawSub == null ? "<null>" : rawSub.GetType().FullName)}");

                        ObjectId rawOid = ObjectId.Null;
                        if (rawSub is ObjectId oid0) rawOid = oid0;
                        else
                        {
                            var po = rawSub?.GetType().GetProperty("ObjectId", BindingFlags.Public | BindingFlags.Instance);
                            if (po != null)
                            {
                                var vv = po.GetValue(rawSub, null);
                                if (vv is ObjectId oid1) rawOid = oid1;
                            }
                        }

                        if (rawOid != ObjectId.Null)
                        {
                            int ridRaw = dlm.FindAcPpRowId(rawOid);
                            Log(ed, $"[UFLOW][DBG] [SubPart {idx}] retry FindAcPpRowId(rawOid)={ridRaw}");
                            if (ridRaw > 0)
                            {
                                hasAnyRowId = true;

                                var pr = SafeGetAllProperties(dlm, rawOid, ridRaw);
                                if (pr != null)
                                {
                                    Log(ed, $"[UFLOW][DBG] [SubPart {idx}] ---- DLM property dump (rawOid rowId={ridRaw}) ----");
                                    DumpAllPropertiesToLog(ed, pr, $"[SubPart {idx}] ");
                                }
                            }
                        }
                    }
                    catch { /* ignore */ }

                    continue;
                }

                hasAnyRowId = true;

                var dict = SafeGetAllProperties(dlm, subKeyObj, rowId);

                Log(ed, $"[UFLOW][DBG] [SubPart {idx}] KeyProps:");

                var lt = FindFirst(dict, new[] { "LineTag", "LineNumber", "LineNo", "LINETAG", "LineNumberTag" });
                var desc = FindFirst(dict, new[] { "PartFamilyLongDesc", "PartFamilyLongDescription", "LongDesc", "LongDescription", "Description", "Desc" });
                var sz = FindFirst(dict, new[] { "Size", "NominalSize", "PartSize", "NominalDiameter", "ND1" });
                var it = FindFirst(dict, new[] { "InstallType", "InstallationType", "Install_Type", "Install" });

                Log(ed, $"[UFLOW][DBG]   LineTag = '{lt.value}' (found={lt.found} key='{lt.keyUsed}')");
                Log(ed, $"[UFLOW][DBG]   PartFamilyLongDesc = '{desc.value}' (found={desc.found} key='{desc.keyUsed}')");
                Log(ed, $"[UFLOW][DBG]   Size = '{sz.value}' (found={sz.found} key='{sz.keyUsed}')");
                Log(ed, $"[UFLOW][DBG]   InstallType = '{it.value}' (found={it.found} key='{it.keyUsed}')");

                if (lt.found) hasLineTag = true;
                if (desc.found) hasDesc = true;
                if (sz.found) hasSize = true;
                if (it.found) hasInstallType = true;

                Log(ed, $"[UFLOW][DBG] [SubPart {idx}] KeySearch (contains):");
                Log(ed, $"[UFLOW][DBG]   'line' -> {FmtKeyList(FindKeysContains(dict, new[] { "line", "tag" }, 20))}");
                Log(ed, $"[UFLOW][DBG]   'desc' -> {FmtKeyList(FindKeysContains(dict, new[] { "desc", "long", "family" }, 20))}");
                Log(ed, $"[UFLOW][DBG]   'size' -> {FmtKeyList(FindKeysContains(dict, new[] { "size", "nominal", "nd" }, 20))}");
                Log(ed, $"[UFLOW][DBG]   'install' -> {FmtKeyList(FindKeysContains(dict, new[] { "install", "type" }, 20))}");

                Log(ed, $"[UFLOW][DBG] [SubPart {idx}] ---- DLM property sample (first 60) ----");
                int c = 0;
                foreach (var kv in dict)
                {
                    Log(ed, $"[UFLOW][DBG]  {kv.Key} = {kv.Value}");
                    c++;
                    if (c >= 60) break;
                }
            }

            Log(ed, "[UFLOW][DBG] Connector SubPart verdict:");
            Log(ed, $"[UFLOW][DBG]   hasAnyRowId={hasAnyRowId}, hasLineTag={hasLineTag}, hasDesc={hasDesc}, hasSize={hasSize}, hasInstallType={hasInstallType}");
        }


        private static int TryFindRowIdFromAcPpObject(DataLinksManager dlm, object acPpObj)
        {
            if (dlm == null || acPpObj == null) return -1;

            try
            {
                if (acPpObj is ObjectId oid)
                {
                    try { return dlm.FindAcPpRowId(oid); } catch { /* ignore */ }
                }

                if (acPpObj is int i)
                {
                    return i;
                }

                // FindAcPpRowId のオーバーロード（引数型が ObjectId 以外）に対応
                var t = dlm.GetType();
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .Where(m => string.Equals(m.Name, "FindAcPpRowId", StringComparison.Ordinal) && m.GetParameters().Length == 1);

                var argType = acPpObj.GetType();
                foreach (var m in methods)
                {
                    var p = m.GetParameters()[0].ParameterType;
                    if (p.IsAssignableFrom(argType))
                    {
                        var v = m.Invoke(dlm, new object[] { acPpObj });
                        if (v is int rid) return rid;
                        if (v != null && int.TryParse(v.ToString(), out var rid2)) return rid2;
                    }
                }

                // acPpObj がラッパー型で ObjectId を持つケース
                var pObjId = argType.GetProperty("ObjectId", BindingFlags.Public | BindingFlags.Instance);
                if (pObjId != null)
                {
                    var v = pObjId.GetValue(acPpObj, null);
                    if (v is ObjectId oid2)
                    {
                        try { return dlm.FindAcPpRowId(oid2); } catch { /* ignore */ }
                    }
                }
            }
            catch
            {
                // ignore
            }

            return -1;
        }

        private static object TryGetAllPropertiesByObject(DataLinksManager dlm, object acPpObj)
        {
            if (dlm == null || acPpObj == null) return null;

            try
            {
                // GetAllProperties のオーバーロード（引数型が ObjectId 以外）に対応
                var t = dlm.GetType();
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .Where(m => string.Equals(m.Name, "GetAllProperties", StringComparison.Ordinal) && m.GetParameters().Length == 2);

                var argType = acPpObj.GetType();

                foreach (var m in methods)
                {
                    var ps = m.GetParameters();
                    var p0 = ps[0].ParameterType;
                    var p1 = ps[1].ParameterType;

                    if (p1 != typeof(bool)) continue;
                    if (!p0.IsAssignableFrom(argType)) continue;

                    return m.Invoke(dlm, new object[] { acPpObj, true });
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static object TryMakeAcPpObjectId(DataLinksManager dlm, ObjectId baseOid, int index)
        {
            try
            {
                // Plant 3D のバージョン/環境で戻り型が異なるため、反射で呼び出す
                var mi = dlm.GetType().GetMethod("MakeAcPpObjectId", BindingFlags.Public | BindingFlags.Instance);
                if (mi != null)
                {
                    return mi.Invoke(dlm, new object[] { baseOid, index });
                }
                return null;
            }
            catch
            {
                try
                {
                    var mi = dlm.GetType().GetMethod("MakeAcPpObjectId", BindingFlags.Public | BindingFlags.Instance);
                    if (mi != null)
                    {
                        var v = mi.Invoke(dlm, new object[] { baseOid, index });
                        if (v is ObjectId oid) return oid;
                    }
                }
                catch { }
                return ObjectId.Null;
            }
        }

        private static (bool found, string value, string keyUsed) FindFirst(Dictionary<string, string> dict, IEnumerable<string> candidateKeys)
        {
            foreach (var k in candidateKeys)
            {
                if (dict.ContainsKey(k))
                {
                    var v = dict[k] ?? "";
                    return (true, v, k);
                }
            }
            return (false, "", "");
        }

        private static List<string> FindKeysContains(Dictionary<string, string> dict, IEnumerable<string> tokensLower, int limit)
        {
            var tokens = tokensLower.Where(t => !string.IsNullOrWhiteSpace(t)).Select(t => t.ToLowerInvariant()).ToList();
            var list = new List<string>();
            foreach (var k in dict.Keys)
            {
                var kl = (k ?? "").ToLowerInvariant();
                bool hit = true;
                foreach (var t in tokens)
                {
                    if (!kl.Contains(t)) { hit = false; break; }
                }
                if (hit)
                {
                    list.Add(k);
                    if (list.Count >= limit) break;
                }
            }
            return list;
        }

        private static string FmtKeyList(List<string> keys)
        {
            if (keys == null || keys.Count == 0) return "(none)";
            return "{ " + string.Join(", ", keys) + " }";
        }
    }
}
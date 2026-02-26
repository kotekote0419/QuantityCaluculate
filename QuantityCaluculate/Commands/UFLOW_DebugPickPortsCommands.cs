// UFLOW_DebugPickPortsCommands_v11_MinProbe_PartPort_PortTable.cs
// Debug command: UFLOW_DEBUG_PICK_PORTS
// v11 (minimal probe):
//  - Write logs to a timestamped text file (Desktop)
//  - Only outputs two probes:
//      (1) DLM relationship probe for Part->Port via GetRelatedRowIds("PartPort","Part",baseRowId,"Port")
//      (2) PnPDatabase Port table probe: dump column names, then try likely FK columns with Select("{col}={baseRowId}")
//
// Notes:
//  - Size parsing is NOT used.
//  - Reflection-heavy to tolerate Plant 3D version/API differences.
//  - If something is not available, it logs the reason and continues.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;

namespace UFLOW
{
    public class UFLOW_DebugPickPortsCommands
    {
        private static StreamWriter _sw;
        private static Editor _ed;

        [CommandMethod("UFLOW_DEBUG_PICK_PORTS")]
        public void DebugPickPorts()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;

            _ed = ed;

            string logPath = null;

            try
            {
                logPath = BuildDesktopLogPath("UFLOW_DEBUG_PICK_PORTS_MinProbe");
                _sw = new StreamWriter(logPath, false, new UTF8Encoding(false));

                ed.WriteMessage("\n[UFLOW][DBG] Minimal probe logging to:");
                ed.WriteMessage("\n[UFLOW][DBG] " + logPath);

                var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick entity to probe (PartPort/PortTable): ");
                peo.AllowNone = false;
                var per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK)
                {
                    Log("[UFLOW][DBG] Cancelled.");
                    return;
                }

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead, false) as Entity;
                    if (ent == null)
                    {
                        Log("[UFLOW][DBG] Selected object is not an Entity.");
                        return;
                    }

                    Log($"[UFLOW][DBG] Picked: Handle={ent.Handle} Type={ent.GetType().FullName}");

                    var dlm = TryGetDataLinksManager(ed);
                    if (dlm == null)
                    {
                        Log("[UFLOW][DBG] DataLinksManager not available.");
                        return;
                    }

                    int baseRowId = TryFindRowIdFromObjectId(dlm, per.ObjectId);
                    Log($"[UFLOW][DBG] BaseObject rowId={baseRowId}");

                    if (baseRowId <= 0)
                    {
                        Log("[UFLOW][DBG] Base rowId not found (<=0).");
                        return;
                    }

                    // quick key props (LineNumberTag / PnPClassName / ConnectionPortCount)
                    var baseProps = TryGetAllPropertiesFlexible(dlm, baseRowId, true);
                    LogKeyProp(baseProps, "PnPClassName");
                    LogKeyProp(baseProps, "PnPGuid");
                    LogKeyProp(baseProps, "LineNumberTag");
                    LogKeyProp(baseProps, "ConnectionPortCount");

                    // (1) PartPort relationship probe
                    ProbePartPortRelationship(dlm, baseRowId);

                    // (2) Port table probe (PnPDatabase -> Port table)
                    ProbePortTable(dlm, baseRowId);

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                SafeEdWrite("\n[UFLOW][DBG] Exception: " + ex.GetType().Name + " / " + ex.Message);
                try { Log("[UFLOW][DBG] Exception: " + ex); } catch { }
            }
            finally
            {
                try
                {
                    if (_sw != null)
                    {
                        _sw.Flush();
                        _sw.Dispose();
                    }
                }
                catch { }
                _sw = null;

                if (!string.IsNullOrWhiteSpace(logPath))
                {
                    SafeEdWrite("\n[UFLOW][DBG] Log saved: " + logPath);
                }

                _ed = null;
            }
        }

        // ----------------------------
        // Probe (1): DLM relationship PartPort
        // ----------------------------
        private static void ProbePartPortRelationship(DataLinksManager dlm, int baseRowId)
        {
            Log("");
            Log("[UFLOW][DBG] ---- Probe#1: GetRelatedRowIds(\"PartPort\",\"Part\", baseRowId, \"Port\") ----");

            var mi = FindGetRelatedRowIds4(dlm);
            if (mi == null)
            {
                Log("[UFLOW][DBG] GetRelatedRowIds(relationshipType,role1,rowId,role2) not found.");
                Log("[UFLOW][DBG] ---- Probe#1 end ----");
                return;
            }

            string rel = "PartPort";
            string role1 = "Part";
            string role2 = "Port";

            var rids = InvokeRelatedRowIds(mi, dlm, rel, role1, baseRowId, role2);
            if (rids.Count == 0)
            {
                // small fallback set (still minimal)
                var variants = new List<(string rel, string r1, string r2)>
                {
                    ("PartPort", "Owner", "Port"),
                    ("PartPort", "Component", "Port"),
                    ("PartPort", "Part", "PartPort"),
                    ("PartPort", "Part", "Ports"),
                };

                foreach (var v in variants)
                {
                    var tmp = InvokeRelatedRowIds(mi, dlm, v.rel, v.r1, baseRowId, v.r2);
                    if (tmp.Count > 0)
                    {
                        Log($"[UFLOW][DBG] Hit with rel='{v.rel}', role1='{v.r1}', role2='{v.r2}' -> {tmp.Count} rowIds");
                        rids = tmp;
                        break;
                    }
                }
            }

            Log($"[UFLOW][DBG] Related Port rowIds count={rids.Count}");
            if (rids.Count > 0)
            {
                Log("[UFLOW][DBG] Related Port rowIds (first 30): " + string.Join(", ", rids.Take(30)));

                // dump key port props for first N
                int n = 0;
                foreach (var rid in rids.Take(10))
                {
                    n++;
                    var p = TryGetAllPropertiesFlexible(dlm, rid, true);
                    string portName = GetProp(p, "PortName");
                    string nd = GetProp(p, "NominalDiameter");
                    string nu = GetProp(p, "NominalUnit");
                    string endType = GetProp(p, "EndType");
                    Log($"[UFLOW][DBG] [PortRid {rid}] PortName='{portName}' NominalDiameter='{nd}' NominalUnit='{nu}' EndType='{endType}'");
                }
                if (rids.Count > 10) Log("[UFLOW][DBG] (port prop dump truncated)");
            }

            Log("[UFLOW][DBG] ---- Probe#1 end ----");
        }

        private static MethodInfo FindGetRelatedRowIds4(DataLinksManager dlm)
        {
            try
            {
                var mis = dlm.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public);
                foreach (var m in mis)
                {
                    if (m.Name != "GetRelatedRowIds") continue;
                    var ps = m.GetParameters();
                    if (ps.Length != 4) continue;
                    if (ps[0].ParameterType != typeof(string)) continue;
                    if (ps[1].ParameterType != typeof(string)) continue;
                    if (ps[2].ParameterType != typeof(int)) continue;
                    if (ps[3].ParameterType != typeof(string)) continue;
                    return m;
                }
            }
            catch { }
            return null;
        }

        private static List<int> InvokeRelatedRowIds(MethodInfo mi, DataLinksManager dlm, string relationshipType, string role1, int rowId, string role2)
        {
            try
            {
                var ret = mi.Invoke(dlm, new object[] { relationshipType, role1, rowId, role2 });
                return CoerceRowIdArray(ret);
            }
            catch (System.Exception ex)
            {
                Log($"[UFLOW][DBG] RelatedRowIds invoke failed: rel='{relationshipType}' role1='{role1}' role2='{role2}' -> {ex.GetType().Name}/{ex.Message}");
                return new List<int>();
            }
        }

        private static List<int> CoerceRowIdArray(object ret)
        {
            var list = new List<int>();
            if (ret == null) return list;

            try
            {
                // most likely IEnumerable<int> or IEnumerable
                if (ret is IEnumerable en)
                {
                    foreach (var x in en)
                    {
                        if (x == null) continue;
                        if (x is int i) { if (i > 0) list.Add(i); continue; }

                        // some Plant 3D types wrap value (RowId property)
                        int v = TryGetIntMember(x, "Value");
                        if (v <= 0) v = TryGetIntMember(x, "RowId");
                        if (v > 0) list.Add(v);
                    }
                }
            }
            catch { }

            return list.Distinct().Where(i => i > 0).ToList();
        }

        // ----------------------------
        // Probe (2): PnPDatabase Port table
        // ----------------------------
        private static void ProbePortTable(DataLinksManager dlm, int baseRowId)
        {
            Log("");
            Log("[UFLOW][DBG] ---- Probe#2: PnPDatabase -> Port table (columns + where candidates) ----");

            object pnpDb = TryGetPnPDatabase(dlm);
            if (pnpDb == null)
            {
                Log("[UFLOW][DBG] PnPDatabase not available via DataLinksManager (public GetPnPDatabase/PnPDatabase).");
                Log("[UFLOW][DBG] ---- Probe#2 end ----");
                return;
            }

            object portTable = TryGetPnPTable(pnpDb, "Port");
            if (portTable == null)
            {
                Log("[UFLOW][DBG] Port table not found in PnPDatabase.");
                Log("[UFLOW][DBG] ---- Probe#2 end ----");
                return;
            }

            // columns
            var columns = TryGetTableColumnNames(portTable);
            Log($"[UFLOW][DBG] Port table columns count={columns.Count}");
            if (columns.Count > 0)
            {
                Log("[UFLOW][DBG] Port table columns (first 120):");
                foreach (var c in columns.Take(120))
                    Log("[UFLOW][DBG]  - " + c);
                if (columns.Count > 120) Log("[UFLOW][DBG]  (columns truncated)");
            }
            else
            {
                Log("[UFLOW][DBG] Port table columns not accessible via reflection (Columns).");
            }

            // where candidates (only if column exists)
            var candidateColsPriority = new[]
            {
                "Owner", "OwnerId", "Parent", "ParentId",
                "Part", "PartId", "Component", "ComponentId",
                "EngineeringItemId", "EngineeringItemsId", "EngItemId", "ItemId",
                "RowOwner", "RowOwnerId", "OwnerRowId", "PartRowId",
            };

            var candidateCols = new List<string>();
            foreach (var c in candidateColsPriority)
            {
                if (columns.Any(x => string.Equals(x, c, StringComparison.OrdinalIgnoreCase)))
                {
                    // keep the actual cased name if possible
                    candidateCols.Add(columns.First(x => string.Equals(x, c, StringComparison.OrdinalIgnoreCase)));
                }
            }

            // heuristic: any column contains "owner"/"part"/"parent" and ends with "id"
            foreach (var c in columns)
            {
                var cl = (c ?? "").ToLowerInvariant();
                if (candidateCols.Any(x => string.Equals(x, c, StringComparison.OrdinalIgnoreCase))) continue;

                bool looksKey = (cl.Contains("owner") || cl.Contains("part") || cl.Contains("parent"))
                                && (cl.EndsWith("id") || cl.EndsWith("rowid"));
                if (looksKey) candidateCols.Add(c);
            }

            candidateCols = candidateCols.Distinct(StringComparer.OrdinalIgnoreCase).Take(20).ToList();
            Log("[UFLOW][DBG] Where candidate columns (max 20): " + (candidateCols.Count == 0 ? "(none)" : string.Join(", ", candidateCols)));

            int anyHit = 0;

            foreach (var col in candidateCols)
            {
                var rows = TrySelectWhere(portTable, $"{col}={baseRowId}");
                int count = rows.Count;
                if (count <= 0) continue;

                anyHit++;
                Log($"[UFLOW][DBG] HIT where '{col}={baseRowId}' -> rows={count}");

                foreach (var row in rows.Take(10))
                {
                    int rid = TryGetRowId(row);
                    string portName = TryGetStringField(row, "PortName");
                    string nd = TryGetStringField(row, "NominalDiameter");
                    string nu = TryGetStringField(row, "NominalUnit");
                    string endType = TryGetStringField(row, "EndType");

                    Log($"[UFLOW][DBG]  PortRow rowId={rid} PortName='{portName}' NominalDiameter='{nd}' NominalUnit='{nu}' EndType='{endType}'");
                }
                if (rows.Count > 10) Log("[UFLOW][DBG]  (rows truncated)");
            }

            if (anyHit == 0)
            {
                Log("[UFLOW][DBG] No hits for where candidates. Next step: inspect Port table columns and add correct FK column name(s).");
            }

            Log("[UFLOW][DBG] ---- Probe#2 end ----");
        }

        // ----------------------------
        // Reflection helpers (PnPDatabase/Table/Select/Row)
        // ----------------------------
        private static object TryGetPnPDatabase(DataLinksManager dlm)
        {
            if (dlm == null) return null;

            // property
            try
            {
                var pi = dlm.GetType().GetProperty("PnPDatabase", BindingFlags.Instance | BindingFlags.Public);
                if (pi != null)
                {
                    var v = pi.GetValue(dlm);
                    if (v != null) return v;
                }
            }
            catch { }

            // method (public)
            foreach (var name in new[] { "GetPnPDatabase", "get_PnPDatabase" })
            {
                try
                {
                    var mi = dlm.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.Public);
                    if (mi == null) continue;
                    var v = mi.Invoke(dlm, null);
                    if (v != null) return v;
                }
                catch { }
            }

            // method (nonpublic) - for Plant 3D version differences
            foreach (var name in new[] { "GetPnPDatabase", "get_PnPDatabase" })
            {
                try
                {
                    var mi = dlm.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic);
                    if (mi == null) continue;
                    var v = mi.Invoke(dlm, null);
                    if (v != null) return v;
                }
                catch { }
            }

            return null;
        }

        private static object TryGetPnPTable(object pnpDb, string tableName)
        {
            if (pnpDb == null || string.IsNullOrWhiteSpace(tableName)) return null;

            // pnpDb.Tables["Port"]
            try
            {
                var piTables = pnpDb.GetType().GetProperty("Tables", BindingFlags.Instance | BindingFlags.Public);
                var tablesObj = piTables?.GetValue(pnpDb);
                if (tablesObj != null)
                {
                    var piItem = tablesObj.GetType().GetProperty("Item", BindingFlags.Instance | BindingFlags.Public);
                    if (piItem != null)
                    {
                        var t = piItem.GetValue(tablesObj, new object[] { tableName });
                        if (t != null) return t;
                    }
                }
            }
            catch { }

            // pnpDb.GetTable("Port")
            foreach (var m in new[] { "GetTable", "get_Table" })
            {
                try
                {
                    var mi = pnpDb.GetType().GetMethod(m, BindingFlags.Instance | BindingFlags.Public, null,
                        new[] { typeof(string) }, null);
                    if (mi == null) continue;
                    var t = mi.Invoke(pnpDb, new object[] { tableName });
                    if (t != null) return t;
                }
                catch { }
            }

            return null;
        }

        private static List<string> TryGetTableColumnNames(object table)
        {
            var list = new List<string>();
            if (table == null) return list;

            try
            {
                var piCols = table.GetType().GetProperty("Columns", BindingFlags.Instance | BindingFlags.Public);
                var colsObj = piCols?.GetValue(table);
                if (colsObj is IEnumerable en)
                {
                    foreach (var c in en)
                    {
                        if (c == null) continue;
                        string name =
                            TryGetStringMember(c, "Name") ??
                            TryGetStringMember(c, "ColumnName") ??
                            TryGetStringMember(c, "FieldName");
                        if (!string.IsNullOrWhiteSpace(name)) list.Add(name);
                    }
                }
            }
            catch { }

            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static List<object> TrySelectWhere(object table, string where)
        {
            var list = new List<object>();
            if (table == null || string.IsNullOrWhiteSpace(where)) return list;

            try
            {
                var mi = table.GetType().GetMethod("Select", BindingFlags.Instance | BindingFlags.Public, null,
                    new[] { typeof(string) }, null);
                if (mi == null) return list;

                var ret = mi.Invoke(table, new object[] { where });
                if (ret is IEnumerable en)
                {
                    foreach (var r in en) if (r != null) list.Add(r);
                }
            }
            catch (System.Exception ex)
            {
                Log($"[UFLOW][DBG] Port.Select(\"{where}\") failed: {ex.GetType().Name}/{ex.Message}");
            }

            return list;
        }

        private static int TryGetRowId(object row)
        {
            if (row == null) return 0;

            try
            {
                var pi = row.GetType().GetProperty("RowId", BindingFlags.Instance | BindingFlags.Public);
                if (pi != null)
                {
                    var v = pi.GetValue(row);
                    if (v is int i) return i;
                }
            }
            catch { }

            // sometimes field name differs
            int a = TryGetIntMember(row, "Id");
            if (a > 0) return a;

            return 0;
        }

        private static string TryGetStringField(object row, string fieldName)
        {
            if (row == null || string.IsNullOrWhiteSpace(fieldName)) return "";

            // indexer row["Field"]
            try
            {
                var piItem = row.GetType().GetProperty("Item", BindingFlags.Instance | BindingFlags.Public);
                if (piItem != null)
                {
                    var idx = piItem.GetIndexParameters();
                    if (idx != null && idx.Length == 1 && idx[0].ParameterType == typeof(string))
                    {
                        var v = piItem.GetValue(row, new object[] { fieldName });
                        return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture);
                    }
                }
            }
            catch { }

            // method GetValue(string)
            foreach (var m in new[] { "GetValue", "get_Item" })
            {
                try
                {
                    var mi = row.GetType().GetMethod(m, BindingFlags.Instance | BindingFlags.Public, null,
                        new[] { typeof(string) }, null);
                    if (mi == null) continue;
                    var v = mi.Invoke(row, new object[] { fieldName });
                    return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture);
                }
                catch { }
            }

            // direct property
            try
            {
                var pi = row.GetType().GetProperty(fieldName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (pi != null)
                {
                    var v = pi.GetValue(row);
                    return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture);
                }
            }
            catch { }

            return "";
        }

        // ----------------------------
        // DLM helpers (RowId / Properties)
        // ----------------------------
        private static DataLinksManager TryGetDataLinksManager(Editor ed)
        {
            try
            {
                var project = PlantApplication.CurrentProject;
                if (project == null) return null;

                var part = project.ProjectParts["Piping"];
                if (part == null) return null;

                var pi = part.GetType().GetProperty("DataLinksManager", BindingFlags.Instance | BindingFlags.Public);
                var dlm = pi?.GetValue(part) as DataLinksManager;
                return dlm;
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage("\n[UFLOW][DBG] TryGetDataLinksManager failed: " + ex.GetType().Name + "/" + ex.Message);
                return null;
            }
        }

        private static int TryFindRowIdFromObjectId(DataLinksManager dlm, ObjectId oid)
        {
            if (dlm == null || oid.IsNull) return 0;

            // FindAcPpRowId(ObjectId) exists
            try
            {
                var mi = dlm.GetType().GetMethod("FindAcPpRowId", BindingFlags.Instance | BindingFlags.Public, null,
                    new[] { typeof(ObjectId) }, null);
                if (mi != null)
                {
                    var v = mi.Invoke(dlm, new object[] { oid });
                    if (v is int i && i > 0) return i;
                }
            }
            catch { }

            // fallback: any method that accepts ObjectId and returns int
            try
            {
                foreach (var m in dlm.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (!string.Equals(m.Name, "FindAcPpRowId", StringComparison.OrdinalIgnoreCase)) continue;
                    var ps = m.GetParameters();
                    if (ps.Length != 1) continue;
                    if (ps[0].ParameterType != typeof(ObjectId)) continue;
                    var v = m.Invoke(dlm, new object[] { oid });
                    if (v is int i && i > 0) return i;
                }
            }
            catch { }

            return 0;
        }

        private static Dictionary<string, string> TryGetAllPropertiesFlexible(DataLinksManager dlm, int rowId, bool includeClassProperties)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (dlm == null || rowId <= 0) return dict;

            object ret = null;
            try
            {
                ret = dlm.GetAllProperties(rowId, includeClassProperties);
            }
            catch
            {
                // ignore
                ret = null;
            }

            if (ret == null) return dict;

            try
            {
                if (ret is IEnumerable en)
                {
                    foreach (var it in en)
                    {
                        if (it == null) continue;

                        // KeyValuePair<string,string>
                        var t = it.GetType();
                        var piK = t.GetProperty("Key");
                        var piV = t.GetProperty("Value");
                        if (piK != null && piV != null)
                        {
                            var k = Convert.ToString(piK.GetValue(it), CultureInfo.InvariantCulture) ?? "";
                            var v = Convert.ToString(piV.GetValue(it), CultureInfo.InvariantCulture) ?? "";
                            if (!string.IsNullOrWhiteSpace(k))
                                dict[k] = v;
                            continue;
                        }

                        // DictionaryEntry
                        if (it is DictionaryEntry de)
                        {
                            var k = Convert.ToString(de.Key, CultureInfo.InvariantCulture) ?? "";
                            var v = Convert.ToString(de.Value, CultureInfo.InvariantCulture) ?? "";
                            if (!string.IsNullOrWhiteSpace(k))
                                dict[k] = v;
                            continue;
                        }
                    }
                }
            }
            catch { }

            return dict;
        }

        private static void LogKeyProp(Dictionary<string, string> dict, string key)
        {
            if (dict == null) return;
            if (string.IsNullOrWhiteSpace(key)) return;

            dict.TryGetValue(key, out string v);
            v = v ?? "";
            Log($"[UFLOW][DBG] {key}='{v}'");
        }

        private static string GetProp(Dictionary<string, string> dict, string key)
        {
            if (dict == null || string.IsNullOrWhiteSpace(key)) return "";
            return dict.TryGetValue(key, out string v) ? (v ?? "") : "";
        }

        // ----------------------------
        // Logging
        // ----------------------------
        private static void Log(string msg)
        {
            try
            {
                if (_sw != null)
                {
                    _sw.WriteLine(msg ?? "");
                    return;
                }
            }
            catch { }

            // fallback to editor
            SafeEdWrite("\n" + (msg ?? ""));
        }

        private static void SafeEdWrite(string msg)
        {
            try { _ed?.WriteMessage(msg); } catch { }
        }

        private static string BuildDesktopLogPath(string prefix)
        {
            var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var name = string.IsNullOrWhiteSpace(prefix) ? "UFLOW_DEBUG" : prefix;
            var file = name + "_" + ts + ".txt";
            var dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            return Path.Combine(dir, file);
        }

        // ----------------------------
        // Small reflection helpers
        // ----------------------------
        private static int TryGetIntMember(object obj, string memberName)
        {
            if (obj == null || string.IsNullOrWhiteSpace(memberName)) return 0;

            try
            {
                var pi = obj.GetType().GetProperty(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (pi != null)
                {
                    var v = pi.GetValue(obj);
                    if (v is int i) return i;
                    if (v != null)
                    {
                        if (int.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), out int parsed))
                            return parsed;
                    }
                }
            }
            catch { }

            try
            {
                var fi = obj.GetType().GetField(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (fi != null)
                {
                    var v = fi.GetValue(obj);
                    if (v is int i) return i;
                    if (v != null)
                    {
                        if (int.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), out int parsed))
                            return parsed;
                    }
                }
            }
            catch { }

            return 0;
        }

        private static string TryGetStringMember(object obj, string memberName)
        {
            if (obj == null || string.IsNullOrWhiteSpace(memberName)) return null;

            try
            {
                var pi = obj.GetType().GetProperty(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (pi != null)
                {
                    var v = pi.GetValue(obj);
                    return v == null ? null : Convert.ToString(v, CultureInfo.InvariantCulture);
                }
            }
            catch { }

            try
            {
                var fi = obj.GetType().GetField(memberName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (fi != null)
                {
                    var v = fi.GetValue(obj);
                    return v == null ? null : Convert.ToString(v, CultureInfo.InvariantCulture);
                }
            }
            catch { }

            return null;
        }
    }
}

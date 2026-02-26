using System;
using System.Collections;
using System.Collections.Generic;
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
using Autodesk.ProcessPower.P3dProjectParts;

namespace UFLOW
{
    public class UFLOW_DebugPartConnectionCommands
    {
        private static StreamWriter _sw;
        private static Editor _ed;

        [CommandMethod("UFLOW_DEBUG_PARTCONNECTION")]
        public void DebugPartConnection()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var ed = doc.Editor;
            var db = doc.Database;
            _ed = ed;

            string logPath = null;

            try
            {
                logPath = BuildDesktopLogPath("UFLOW_DEBUG_PARTCONNECTION");
                _sw = new StreamWriter(logPath, false, new UTF8Encoding(false));

                ed.WriteMessage("\n[UFLOW][DBG] P3dPartConnection probe logging to:\n" + logPath);

                var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick entity to probe (P3dPartConnection): ");
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

                    var dlm = TryGetDataLinksManager();
                    if (dlm == null)
                    {
                        Log("[UFLOW][DBG] DataLinksManager not available.\n" +
                            "[UFLOW][DBG] Make sure this is a Piping model, and ProjectParts['Piping'] is present.");
                        return;
                    }

                    int baseRowId = TryFindRowIdFromObjectId(dlm, per.ObjectId);
                    Log($"[UFLOW][DBG] BaseObject rowId={baseRowId}");
                    if (baseRowId <= 0)
                    {
                        Log("[UFLOW][DBG] Base rowId not found (<=0).\n");
                        return;
                    }

                    ProbeP3dPartConnection(dlm, baseRowId);
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
                try { _sw?.Flush(); _sw?.Dispose(); } catch { }
                _sw = null;

                if (!string.IsNullOrWhiteSpace(logPath))
                    SafeEdWrite("\n[UFLOW][DBG] Log saved: " + logPath);

                _ed = null;
            }
        }

        // --- DataLinksManager retrieval (no foreach over ProjectParts) ---
        private static DataLinksManager TryGetDataLinksManager()
        {
            try
            {
                var plantProj = PlantApplication.CurrentProject;
                if (plantProj == null) return null;

                // Most common: ProjectParts["Piping"] is a PipingProject
                object parts = plantProj.ProjectParts;
                if (parts == null) return null;

                // Try indexer: parts["Piping"]
                var idx = parts.GetType().GetProperty("Item", new[] { typeof(string) });
                if (idx != null)
                {
                    foreach (var key in new[] { "Piping", "PIPING" })
                    {
                        try
                        {
                            var v = idx.GetValue(parts, new object[] { key });
                            if (v is PipingProject pp) return pp.DataLinksManager;
                        }
                        catch { }
                    }
                }

                // As a fallback, attempt to find a DataLinksManager property on any ProjectPart-like objects
                // without assuming dictionary entry types.
                if (parts is IEnumerable en)
                {
                    foreach (var it in en)
                    {
                        if (it == null) continue;

                        // If iterator yields a PipingProject directly
                        if (it is PipingProject pp) return pp.DataLinksManager;

                        // If iterator yields a key/value container, try to read Value property
                        var vt = it.GetType();
                        var valProp = vt.GetProperty("Value", BindingFlags.Instance | BindingFlags.Public);
                        if (valProp != null)
                        {
                            var val = valProp.GetValue(it);
                            if (val is PipingProject pp2) return pp2.DataLinksManager;
                        }
                    }
                }

                return null;
            }
            catch (System.Exception ex)
            {
                Log("[UFLOW][DBG] TryGetDataLinksManager failed: " + ex.Message);
                return null;
            }
        }

        private static void ProbeP3dPartConnection(DataLinksManager dlm, int baseRowId)
        {
            Log("\n[UFLOW][DBG] ---- Probe: P3dPartConnection ----");

            var mi = FindGetRelatedRowIds4(dlm);
            if (mi == null)
            {
                Log("[UFLOW][DBG] GetRelatedRowIds(string,string,int,string) not found on DataLinksManager.");
                return;
            }

            var relCandidates = new List<string> { "P3dPartConnection", "P3dPartConnections", "PartConnection", "PartConnections" };
            var rolePairs = new List<(string fromRole, string toRole)>
            {
                ("Part1", "Part2"),
                ("Part2", "Part1"),
                ("Part", "ConnectedPart"),
                ("Owner", "Connected"),
            };

            bool anyHit = false;
            foreach (var rel in relCandidates)
            {
                foreach (var (fromRole, toRole) in rolePairs)
                {
                    var rids = InvokeRelatedRowIds(mi, dlm, rel, fromRole, baseRowId, toRole);
                    if (rids.Count == 0) continue;

                    anyHit = true;
                    Log($"[UFLOW][DBG] HIT {rel}:{fromRole}->{toRole} : {rids.Count} | {string.Join(", ", rids.Take(50))}");
                }
            }

            if (!anyHit)
            {
                Log("[UFLOW][DBG] No related rowIds found via P3dPartConnection candidates.");
                Log("[UFLOW][DBG] If this remains zero, we may need a PnPDatabase table-direct probe for P3dPartConnection.");
            }

            Log("[UFLOW][DBG] ---- Probe end ----");
        }

        private static int TryFindRowIdFromObjectId(DataLinksManager dlm, ObjectId oid)
        {
            try
            {
                var t = dlm.GetType();

                var m1 = t.GetMethod("FindAcPpRowId", BindingFlags.Instance | BindingFlags.Public, null, new[] { typeof(ObjectId) }, null);
                if (m1 != null) return ToInt(m1.Invoke(dlm, new object[] { oid }));

                var m2 = t.GetMethod("FindRowId", BindingFlags.Instance | BindingFlags.Public, null, new[] { typeof(ObjectId) }, null);
                if (m2 != null) return ToInt(m2.Invoke(dlm, new object[] { oid }));

                foreach (var m in t.GetMethods(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (m.Name.IndexOf("Find", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    if (m.Name.IndexOf("RowId", StringComparison.OrdinalIgnoreCase) < 0) continue;
                    var ps = m.GetParameters();
                    if (ps.Length == 1 && ps[0].ParameterType == typeof(ObjectId))
                        return ToInt(m.Invoke(dlm, new object[] { oid }));
                }
            }
            catch (System.Exception ex)
            {
                Log("[UFLOW][DBG] TryFindRowIdFromObjectId failed: " + ex.Message);
            }
            return 0;
        }

        private static MethodInfo FindGetRelatedRowIds4(DataLinksManager dlm)
        {
            try
            {
                foreach (var m in dlm.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (m.Name != "GetRelatedRowIds") continue;
                    var ps = m.GetParameters();
                    if (ps.Length == 4 &&
                        ps[0].ParameterType == typeof(string) &&
                        ps[1].ParameterType == typeof(string) &&
                        ps[2].ParameterType == typeof(int) &&
                        ps[3].ParameterType == typeof(string))
                        return m;
                }
            }
            catch { }
            return null;
        }

        private static List<int> InvokeRelatedRowIds(MethodInfo mi, DataLinksManager dlm, string rel, string role1, int rowId, string role2)
        {
            var rids = new List<int>();
            try
            {
                object r = mi.Invoke(dlm, new object[] { rel, role1, rowId, role2 });
                if (r is IEnumerable en)
                {
                    foreach (var it in en)
                    {
                        int v = ToInt(it);
                        if (v > 0) rids.Add(v);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Log($"[UFLOW][DBG] InvokeRelatedRowIds failed rel='{rel}' {role1}->{role2}: {ex.Message}");
            }
            return rids;
        }

        private static int ToInt(object x)
        {
            if (x == null) return 0;
            if (x is int i) return i;
            if (x is long l) return (int)l;
            if (x is short s) return s;
            if (x is string str && int.TryParse(str, out var j)) return j;
            try { return Convert.ToInt32(x); } catch { return 0; }
        }

        private static void Log(string s)
        {
            try { _sw?.WriteLine(s); } catch { }
        }

        private static void SafeEdWrite(string s)
        {
            try { _ed?.WriteMessage(s); } catch { }
        }

        private static string BuildDesktopLogPath(string prefix)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return Path.Combine(desktop, $"{prefix}_{ts}.txt");
        }
    }
}

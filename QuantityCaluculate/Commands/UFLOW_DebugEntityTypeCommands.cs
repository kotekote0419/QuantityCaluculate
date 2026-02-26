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
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFLOW
{
    // EntityType(=GetType().Name) の分布を棚卸しして txt へ出力するデバッグコマンド
    public class UFLOW_DebugEntityTypeCommands
    {
        private static StreamWriter _sw;
        private static Editor _ed;

        [CommandMethod("UFLOW_DEBUG_ENTITYTYPES")]
        public void DebugEntityTypes()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            _ed = doc.Editor;
            var db = doc.Database;

            string logPath = null;

            try
            {
                logPath = BuildDesktopLogPath("UFLOW_DEBUG_ENTITYTYPES");
                _sw = new StreamWriter(logPath, false, new UTF8Encoding(false));

                _ed.WriteMessage("\n[UFLOW][DBG] EntityType inventory logging to:\n" + logPath);

                var dlm = TryGetDataLinksManager();

                int ignoredAutoCad = 0;
                int totalTargets = 0;
                int totalP3dOthers = 0;
                int totalConnectors = 0;

                var targets = new Dictionary<string, Bucket>(StringComparer.Ordinal);
                var p3dOthers = new Dictionary<string, Bucket>(StringComparer.Ordinal);
                var connectors = new Dictionary<string, Bucket>(StringComparer.Ordinal);

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId oid in ms)
                    {
                        if (oid.IsNull || oid.IsErased) continue;

                        var ent = tr.GetObject(oid, OpenMode.ForRead, false) as Entity;
                        if (ent == null) continue;

                        string typeName = ent.GetType().Name;
                        string typeFull = ent.GetType().FullName ?? typeName;
                        string ns = ent.GetType().Namespace ?? "";

                        bool isTarget = (ent is Pipe) || (ent is Part) || (ent is PipeInlineAsset);

                        // ★ P3dConnector を型参照しない（CS0246回避）：名前で判定
                        bool isConnector =
                            string.Equals(ent.GetType().Name, "P3dConnector", StringComparison.Ordinal) ||
                            string.Equals(ent.GetType().FullName, "Autodesk.ProcessPower.PnP3dObjects.P3dConnector", StringComparison.Ordinal);

                        bool isPlant3dLike =
                            isTarget || isConnector ||
                            ns.StartsWith("Autodesk.ProcessPower", StringComparison.OrdinalIgnoreCase) ||
                            typeFull.StartsWith("Autodesk.ProcessPower", StringComparison.OrdinalIgnoreCase);

                        if (!isPlant3dLike)
                        {
                            ignoredAutoCad++;
                            continue;
                        }

                        int rowId = (dlm != null) ? SafeFindRowId(dlm, oid) : 0;

                        if (isConnector)
                        {
                            totalConnectors++;
                            Add(connectors, typeName, typeFull, ent.Handle.ToString(), rowId);
                        }
                        else if (isTarget)
                        {
                            totalTargets++;
                            Add(targets, typeName, typeFull, ent.Handle.ToString(), rowId);
                        }
                        else
                        {
                            totalP3dOthers++;
                            Add(p3dOthers, typeName, typeFull, ent.Handle.ToString(), rowId);
                        }
                    }

                    tr.Commit();
                }

                Log($"[UFLOW][DBG] UFLOW_DEBUG_ENTITYTYPES");
                Log($"[UFLOW][DBG] TotalTargets(Pipe/Part/PipeInlineAsset)={totalTargets}");
                Log($"[UFLOW][DBG] TotalP3dOthers(Autodesk.ProcessPower.* etc)={totalP3dOthers}");
                Log($"[UFLOW][DBG] TotalConnectors(P3dConnector)={totalConnectors}  ※集計対象外だが参考");
                Log($"[UFLOW][DBG] IgnoredAutoCADEntities={ignoredAutoCad}");
                Log("");

                LogSection("[UFLOW][ENTITYTYPE][TARGETS]", targets);
                LogSection("[UFLOW][ENTITYTYPE][P3D_OTHERS]", p3dOthers);
                LogSection("[UFLOW][ENTITYTYPE][CONNECTORS]", connectors);

                Log("");
                Log("[UFLOW][DBG] Notes:");
                Log("[UFLOW][DBG] - EntityType は ent.GetType().Name を使用。");
                Log("[UFLOW][DBG] - TARGETS は現在の材料集計の主対象。P3dConnector は従来通り除外。");
            }
            catch (System.Exception ex)
            {
                try
                {
                    _ed?.WriteMessage("\n[UFLOW][DBG] Exception: " + ex.GetType().Name + " / " + ex.Message);
                    Log("[UFLOW][DBG] Exception: " + ex);
                }
                catch { }
            }
            finally
            {
                try { _sw?.Flush(); _sw?.Dispose(); } catch { }
                _sw = null;

                if (!string.IsNullOrWhiteSpace(logPath))
                {
                    try { _ed?.WriteMessage("\n[UFLOW][DBG] Log saved: " + logPath); } catch { }
                }

                _ed = null;
            }
        }

        private class Bucket
        {
            public string TypeName;
            public string TypeFullName;
            public string ExampleHandle;
            public int ExampleRowId;
            public int Count;
        }

        private static void Add(Dictionary<string, Bucket> dict, string typeName, string typeFull, string handle, int rowId)
        {
            if (!dict.TryGetValue(typeName, out var b))
            {
                b = new Bucket
                {
                    TypeName = typeName,
                    TypeFullName = typeFull,
                    ExampleHandle = handle,
                    ExampleRowId = rowId
                };
                dict[typeName] = b;
            }

            b.Count++;

            if (b.ExampleRowId <= 0 && rowId > 0)
            {
                b.ExampleRowId = rowId;
                b.ExampleHandle = handle;
                b.TypeFullName = typeFull;
            }
        }

        private static void LogSection(string title, Dictionary<string, Bucket> dict)
        {
            Log(title);

            if (dict == null || dict.Count == 0)
            {
                Log("  (none)");
                Log("");
                return;
            }

            foreach (var b in dict.Values
                                 .OrderByDescending(x => x.Count)
                                 .ThenBy(x => x.TypeName, StringComparer.Ordinal))
            {
                Log($"  {b.TypeName}\tcount={b.Count}\texampleHandle={b.ExampleHandle}\trowId={b.ExampleRowId}\tfull={b.TypeFullName}");
            }

            Log("");
        }

        private static void Log(string s)
        {
            try { _sw?.WriteLine(s); } catch { }
        }

        private static string BuildDesktopLogPath(string prefix)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return Path.Combine(desktop, $"{prefix}_{ts}.txt");
        }

        private static DataLinksManager TryGetDataLinksManager()
        {
            try
            {
                var plantProj = PlantApplication.CurrentProject;
                if (plantProj == null) return null;

                object parts = plantProj.ProjectParts;
                if (parts == null) return null;

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

                if (parts is IEnumerable en)
                {
                    foreach (var it in en)
                    {
                        if (it == null) continue;

                        if (it is PipingProject pp) return pp.DataLinksManager;

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
            catch
            {
                return null;
            }
        }

        private static int SafeFindRowId(DataLinksManager dlm, ObjectId oid)
        {
            if (dlm == null || oid.IsNull) return 0;

            try
            {
                var t = dlm.GetType();
                var mi = t.GetMethod("FindRowId", BindingFlags.Instance | BindingFlags.Public, null, new[] { typeof(ObjectId) }, null);
                if (mi != null)
                {
                    object r = mi.Invoke(dlm, new object[] { oid });
                    return ToInt(r);
                }
            }
            catch { }

            return 0;
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
    }
}

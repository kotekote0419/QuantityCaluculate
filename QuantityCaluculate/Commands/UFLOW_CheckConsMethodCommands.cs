using System;
using System.Collections.Generic;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.PnP3dObjects;

namespace UFLOW
{
    /// <summary>
    /// 施工方法（"施工方法"プロパティ）に応じて一時的に色分けして確認する。
    /// ※表示スタイルの切替は行わない（事前に手動で2Dワイヤフレームへ切替）
    /// </summary>
    public class UFLOW_CheckConsMethodCommands
    {
        private const string SphereLayerName = "UFLOW_CHECK_CONSMETHOD";
        private const double SphereRadius = 100.0;

        [CommandMethod("UFLOW_CHECK_CONSMETHOD", CommandFlags.Modal)]
        public void UFLOW_CHECK_CONSMETHOD()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var db = doc.Database;
            var ed = doc.Editor;

            var opt = new PromptKeywordOptions("\nFilter [All(A)/Pipe(P)/Gasket(G)/Buttweld(B)] <All>: ", "All Pipe Gasket Buttweld");
            opt.AllowNone = true;
            opt.Keywords.Default = "All";
            var kres = ed.GetKeywords(opt);
            if (kres.Status != PromptStatus.OK && kres.Status != PromptStatus.None) return;

            string filter = (kres.Status == PromptStatus.None || string.IsNullOrWhiteSpace(kres.StringResult))
                ? "All" : kres.StringResult;

            bool doPipe = filter.Equals("All", StringComparison.OrdinalIgnoreCase) || filter.Equals("Pipe", StringComparison.OrdinalIgnoreCase);
            bool doGasket = filter.Equals("All", StringComparison.OrdinalIgnoreCase) || filter.Equals("Gasket", StringComparison.OrdinalIgnoreCase);
            bool doButtweld = filter.Equals("All", StringComparison.OrdinalIgnoreCase) || filter.Equals("Buttweld", StringComparison.OrdinalIgnoreCase);

            Color acColorOver = Color.FromRgb(0, 176, 240);   // 架設
            Color acColorBury = Color.FromRgb(230, 0, 0);     // 埋設
            Color acColorOthr = Color.FromRgb(180, 0, 255);   // その他

            var originalColors = new Dictionary<ObjectId, Color>();
            var sphereIds = new List<ObjectId>();

            using (doc.LockDocument())
            {
                try
                {
                    EnsureLayer(db, SphereLayerName);

                    int pipeCount = 0;
                    int inlineCount = 0;
                    int gasketSphereCount = 0;
                    int buttweldSphereCount = 0;

                    using (var tr = db.TransactionManager.StartTransaction())
                    {
                        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                        foreach (ObjectId oid in ms)
                        {
                            if (oid.IsNull || oid.IsErased) continue;

                            var ent = tr.GetObject(oid, OpenMode.ForWrite, false) as Entity;
                            if (ent == null) continue;

                            string tn = ent.GetType().Name;

                            if (doPipe && (tn == "Pipe" || tn == "PipeInlineAsset"))
                            {
                                string cm = GetConsMethodFromPart(ent);

                                if (!originalColors.ContainsKey(oid))
                                    originalColors[oid] = ent.Color;

                                ent.Color = PickColor(cm, acColorOver, acColorBury, acColorOthr);

                                if (tn == "Pipe") pipeCount++;
                                else inlineCount++;
                                continue;
                            }

                            if ((doGasket || doButtweld) && tn == "Connector")
                            {
                                var conn = ent as Connector;
                                if (conn == null || conn.AllSubParts == null || conn.AllSubParts.Count == 0) continue;

                                foreach (SubPart sp in conn.AllSubParts)
                                {
                                    if (sp == null) continue;

                                    string spType = sp.PartSizeProperties?.Type ?? "";
                                    bool isG = doGasket && spType.Equals("Gasket", StringComparison.OrdinalIgnoreCase);
                                    bool isB = doButtweld && spType.Equals("Buttweld", StringComparison.OrdinalIgnoreCase);
                                    if (!isG && !isB) continue;

                                    string cm = GetConsMethodFromSubPart(sp);

                                    var sphere = new Solid3d();
                                    sphere.CreateSphere(SphereRadius);
                                    sphere.TransformBy(Matrix3d.Displacement(conn.Position - Point3d.Origin));
                                    sphere.Layer = SphereLayerName;
                                    sphere.Color = PickColor(cm, acColorOver, acColorBury, acColorOthr);

                                    ms.AppendEntity(sphere);
                                    tr.AddNewlyCreatedDBObject(sphere, true);
                                    sphereIds.Add(sphere.ObjectId);

                                    if (isG) gasketSphereCount++;
                                    if (isB) buttweldSphereCount++;
                                }
                            }
                        }

                        tr.Commit();
                    }

                    ed.Regen();
                    try { Application.UpdateScreen(); } catch { }

                    ed.WriteMessage($"\n[UFLOW][CHKCONS] Apply done: Pipe={pipeCount}, PipeInlineAsset={inlineCount}, GasketSphere={gasketSphereCount}, ButtweldSphere={buttweldSphereCount}");
                    ed.WriteMessage("\n[UFLOW] 施工方法の色分けを適用しました。Enterで復元して終了（Escでも復元）。");

                    var wait = new PromptStringOptions("\n(Enter) ") { AllowSpaces = true };
                    ed.GetString(wait);
                }
                finally
                {
                    try
                    {
                        using (var tr2 = db.TransactionManager.StartTransaction())
                        {
                            foreach (var kv in originalColors)
                            {
                                if (kv.Key.IsNull || kv.Key.IsErased) continue;
                                var ent = tr2.GetObject(kv.Key, OpenMode.ForWrite, false) as Entity;
                                if (ent != null) ent.Color = kv.Value;
                            }

                            foreach (var sid in sphereIds)
                            {
                                if (sid.IsNull || sid.IsErased) continue;
                                var ent = tr2.GetObject(sid, OpenMode.ForWrite, false) as Entity;
                                if (ent != null) ent.Erase();
                            }

                            tr2.Commit();
                        }

                        ed.Regen();
                        try { Application.UpdateScreen(); } catch { }

                        ed.WriteMessage($"\n[UFLOW] 復元しました。color={originalColors.Count}, sphere={sphereIds.Count}\n");
                    }
                    catch (System.Exception ex2)
                    {
                        ed.WriteMessage("\n[UFLOW] 復元中に例外: " + ex2.Message + "\n");
                    }
                }
            }
        }

        private static Color PickColor(string consMethod, Color over, Color bury, Color othr)
        {
            if (string.Equals(consMethod, "架設", StringComparison.Ordinal)) return over;
            if (string.Equals(consMethod, "埋設", StringComparison.Ordinal)) return bury;
            return othr;
        }

        private static string GetConsMethodFromPart(Entity ent)
        {
            try
            {
                if (ent == null) return "";
                dynamic d = ent;
                object v = null;
                try { v = d.PartSizeProperties.PropValue("施工方法"); } catch { v = null; }
                return v != null ? v.ToString() : "";
            }
            catch { return ""; }
        }

        private static string GetConsMethodFromSubPart(SubPart sp)
        {
            try
            {
                if (sp?.PartSizeProperties == null) return "";
                if (sp.PartSizeProperties.PropNames != null && sp.PartSizeProperties.PropNames.Contains("施工方法"))
                {
                    var v = sp.PartSizeProperties.PropValue("施工方法");
                    return v != null ? v.ToString() : "";
                }
                return "";
            }
            catch { return ""; }
        }

        private static void EnsureLayer(Database db, string layerName)
        {
            if (db == null) return;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    var ltr = new LayerTableRecord { Name = layerName };
                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }
                tr.Commit();
            }
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

// Plant 3D
using Autodesk.ProcessPower.DataLinks;
using UFlowPlant3D.Services;

namespace UFlow // ←ここはあなたのプロジェクトに合わせて変更
{
    public class DebugFastenerCommands
    {
        /// <summary>
        /// 1つ選択して、rowId/型/PnPClassName/Ports(S1,S2)/距離/ FastenerRowsとの一致を全部ログ出し
        /// </summary>
        [CommandMethod("UFLOW_DEBUG_PICK_PORTS")]
        public void UFLOW_DEBUG_PICK_PORTS()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            DataLinksManager dlm = null;
            try
            {
                dlm = DataLinksManager.GetManager(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] DataLinksManager.GetManager failed: {ex.Message}");
                return;
            }

            // FastenerRowIds（PnPDatabase: Fasteners/Gasket/BoltSet）
            HashSet<int> fastenerRows = new HashSet<int>();
            try
            {
                var list = FastenerCollector.CollectFastenerRowIds(dlm, ed);
                fastenerRows = new HashSet<int>(list);
                ed.WriteMessage($"\n[UFLOW][DBG] FastenerRows(PnPDB) count={fastenerRows.Count}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] FastenerCollector failed (continue): {ex.Message}");
            }

            // Pick
            var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick entity (Pipe/Part/Connector/Fastener) : ");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null)
                {
                    ed.WriteMessage("\n[UFLOW][DBG] Not an Entity.");
                    return;
                }

                string handle = ent.Handle.ToString();
                string typeName = ent.GetType().FullName ?? ent.GetType().Name;

                int rowId = -1;
                try
                {
                    rowId = dlm.FindAcPpRowId(per.ObjectId);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[UFLOW][DBG] FindAcPpRowId failed: {ex.Message}");
                }

                string pnpClass = "";
                try
                {
                    // ObjectIdから取得（あなたのPlantProp.cs前提）
                    pnpClass = PlantProp.GetString(dlm, per.ObjectId, "PnPClassName") ?? "";
                }
                catch { /* ignore */ }

                ed.WriteMessage(
                    $"\n[UFLOW][DBG] Handle={handle} Type={typeName} RowId={rowId} PnPClassName='{pnpClass}'");

                // RowIdがPnPDB FastenerRowsと一致しているか？
                if (rowId > 0)
                {
                    bool inFastenerRows = fastenerRows.Contains(rowId);
                    ed.WriteMessage($"\n[UFLOW][DBG] RowId in FastenerRows(PnPDB)? {inFastenerRows}");
                }

                // rowIdのLineTag等（取れる範囲で）
                if (rowId > 0)
                {
                    try
                    {
                        string lt = (PlantProp.GetString(dlm, rowId, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号") ?? "").Trim();
                        string spec = (PlantProp.GetString(dlm, rowId, "Spec", "SPEC") ?? "").Trim();
                        ed.WriteMessage($"\n[UFLOW][DBG] RowProps: LineTag='{lt}', Spec='{spec}'");
                    }
                    catch
                    {
                        // PlantPropに rowId overload が無い場合でもコンパイルは通らないので注意
                        ed.WriteMessage("\n[UFLOW][DBG] RowProps: (PlantProp rowId overload not available or failed)");
                    }
                }

                // Ports取得（Reflection）
                var ports = GetPortsByReflection(ent);

                ed.WriteMessage($"\n[UFLOW][DBG] Ports count={ports.Count}");
                for (int i = 0; i < ports.Count; i++)
                {
                    string name = ports[i].Name ?? $"(idx:{i})";
                    ed.WriteMessage($"\n[UFLOW][DBG]  Port[{i}] Name='{name}' Pos={Fmt(ports[i].Pos)}");
                }

                // S1/S2 距離
                if (ports.Count >= 2)
                {
                    int iS1 = ports.FindIndex(p => EqPortName(p.Name, "S1"));
                    int iS2 = ports.FindIndex(p => EqPortName(p.Name, "S2"));

                    Point3d p1, p2;
                    string mode;
                    if (iS1 >= 0 && iS2 >= 0)
                    {
                        p1 = ports[iS1].Pos;
                        p2 = ports[iS2].Pos;
                        mode = "S1-S2";
                    }
                    else
                    {
                        p1 = ports[0].Pos;
                        p2 = ports[1].Pos;
                        mode = "Port0-Port1";
                    }

                    double dist = p1.DistanceTo(p2);
                    ed.WriteMessage($"\n[UFLOW][DBG] Thickness({mode}) = {dist} (drawing unit)");
                }
                else
                {
                    ed.WriteMessage("\n[UFLOW][DBG] Not enough ports to compute thickness.");
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// ModelSpace を走査して「FindAcPpRowIdで取れたrowId」がFastenerRows(PnPDB)に一致する数を統計で出す
        /// （一致が0なら、FastenerRowと図形rowIdが別テーブルでズレている可能性が濃厚）
        /// </summary>
        [CommandMethod("UFLOW_DEBUG_FASTENER_ROWID_MATCH")]
        public void UFLOW_DEBUG_FASTENER_ROWID_MATCH()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            DataLinksManager dlm = null;
            try
            {
                dlm = DataLinksManager.GetManager(db);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] DataLinksManager.GetManager failed: {ex.Message}");
                return;
            }

            HashSet<int> fastenerRows = new HashSet<int>();
            try
            {
                fastenerRows = new HashSet<int>(FastenerCollector.CollectFastenerRowIds(dlm, ed));
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] FastenerCollector failed: {ex.Message}");
                return;
            }

            int totalEnt = 0;
            int hasRowId = 0;
            int matched = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId oid in ms)
                {
                    totalEnt++;
                    int rid = -1;
                    try
                    {
                        rid = dlm.FindAcPpRowId(oid);
                    }
                    catch
                    {
                        continue;
                    }

                    if (rid > 0)
                    {
                        hasRowId++;
                        if (fastenerRows.Contains(rid)) matched++;
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n[UFLOW][DBG] ModelSpace total={totalEnt}, hasRowId={hasRowId}, matchedFastenerRowId={matched}, fastenerRowsCount={fastenerRows.Count}");
            ed.WriteMessage("\n[UFLOW][DBG] matched=0 の場合：図形rowIdがFasteners/Gasket/BoltSetのrowIdではない（別テーブルのrowId）可能性が高いです。");
        }

        // -----------------------------
        // Reflection-based port reading
        // -----------------------------

        private class PortInfo
        {
            public string Name;
            public Point3d Pos;
        }

        private static List<PortInfo> GetPortsByReflection(Entity ent)
        {
            var result = new List<PortInfo>();
            if (ent == null) return result;

            object portCollection = null;

            // 1) method GetPorts() / GetPorts(x)
            var t = ent.GetType();
            var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance);

            MethodInfo mi = methods.FirstOrDefault(m => m.Name == "GetPorts" && m.GetParameters().Length == 0);
            if (mi != null)
            {
                try { portCollection = mi.Invoke(ent, null); } catch { /*ignore*/ }
            }
            else
            {
                // GetPorts(PortType) の可能性
                mi = methods.FirstOrDefault(m => m.Name == "GetPorts" && m.GetParameters().Length == 1);
                if (mi != null)
                {
                    try
                    {
                        var pType = mi.GetParameters()[0].ParameterType;
                        object arg = CreateEnumArgIfPossible(pType, "Static") ?? CreateDefaultValue(pType);
                        portCollection = mi.Invoke(ent, new[] { arg });
                    }
                    catch { /*ignore*/ }
                }
            }

            // 2) property Ports / PortCollection
            if (portCollection == null)
            {
                var piPorts = t.GetProperty("Ports", BindingFlags.Public | BindingFlags.Instance);
                if (piPorts != null)
                {
                    try { portCollection = piPorts.GetValue(ent, null); } catch { /*ignore*/ }
                }
            }

            if (portCollection == null)
            {
                var pi = t.GetProperty("PortCollection", BindingFlags.Public | BindingFlags.Instance);
                if (pi != null)
                {
                    try { portCollection = pi.GetValue(ent, null); } catch { /*ignore*/ }
                }
            }

            // iterate
            if (portCollection is IEnumerable en)
            {
                foreach (var p in en)
                {
                    if (p == null) continue;

                    string name = TryGetStringProp(p, "Name") ?? TryGetStringProp(p, "PortName");

                    Point3d pos = TryGetPoint3dProp(p, "Position")
                                  ?? TryGetPoint3dProp(p, "Location")
                                  ?? TryGetPoint3dProp(p, "Point")
                                  ?? Point3d.Origin;

                    result.Add(new PortInfo { Name = name, Pos = pos });
                }
            }

            return result;
        }

        private static object CreateEnumArgIfPossible(Type enumType, string name)
        {
            try
            {
                if (enumType != null && enumType.IsEnum)
                {
                    // "Static" が無い場合もあるので例外は握りつぶし
                    return Enum.Parse(enumType, name, ignoreCase: true);
                }
            }
            catch { }
            return null;
        }

        private static object CreateDefaultValue(Type t)
        {
            try
            {
                if (t == null) return null;
                if (t.IsValueType) return Activator.CreateInstance(t);
            }
            catch { }
            return null;
        }

        private static string TryGetStringProp(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) return null;
                var v = pi.GetValue(obj, null);
                return v?.ToString();
            }
            catch { return null; }
        }

        private static Point3d? TryGetPoint3dProp(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) return null;

                var v = pi.GetValue(obj, null);
                if (v == null) return null;

                if (v is Point3d p3) return p3;

                // X,Y,Z を持つオブジェクトもあるのでフォールバック
                double? x = TryGetDoubleProp(v, "X");
                double? y = TryGetDoubleProp(v, "Y");
                double? z = TryGetDoubleProp(v, "Z");
                if (x.HasValue && y.HasValue && z.HasValue) return new Point3d(x.Value, y.Value, z.Value);
            }
            catch { }
            return null;
        }

        private static double? TryGetDoubleProp(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) return null;
                var v = pi.GetValue(obj, null);
                if (v == null) return null;
                if (v is double d) return d;
                if (double.TryParse(v.ToString(), out var dd)) return dd;
            }
            catch { }
            return null;
        }

        private static bool EqPortName(string a, string b)
        {
            if (a == null || b == null) return false;
            return string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static string Fmt(Point3d p)
        {
            return $"({p.X:0.###},{p.Y:0.###},{p.Z:0.###})";
        }
    }
}

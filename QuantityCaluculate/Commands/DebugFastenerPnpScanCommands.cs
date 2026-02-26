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

using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;

namespace UFlow
{
    public class DebugFastenerPnpScanCommands
    {
        [CommandMethod("UFLOW_DEBUG_SCAN_FASTENERS_LINK")]
        public void UFLOW_DEBUG_SCAN_FASTENERS_LINK()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            DataLinksManager dlm;
            try { dlm = DataLinksManager.GetManager(db); }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] DLM get failed: {ex.Message}");
                return;
            }

            // pick connector
            var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick Connector (gasket-like): ");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            int connRowId = -1;
            try { connRowId = dlm.FindAcPpRowId(per.ObjectId); }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] FindAcPpRowId failed: {ex.Message}");
                return;
            }

            // connector props (via DLM)
            var connProps = SafeGetAllProps(dlm, connRowId);
            string connGuid = GetFirstString(connProps, "PnPGuid", "RowGuid", "Guid", "GUID");

            ed.WriteMessage($"\n[UFLOW][DBG] Connector RowId={connRowId}, PnPGuid='{connGuid}'");

            // thickness
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var ent = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                if (ent != null)
                {
                    var ports = GetPortsByReflection(ent);
                    if (ports.Count >= 2)
                    {
                        var s1 = ports.FirstOrDefault(p => Eq(p.Name, "S1"));
                        var s2 = ports.FirstOrDefault(p => Eq(p.Name, "S2"));
                        Point3d p1 = (s1 != null) ? s1.Pos : ports[0].Pos;
                        Point3d p2 = (s2 != null) ? s2.Pos : ports[1].Pos;
                        ed.WriteMessage($"\n[UFLOW][DBG] Thickness(S1-S2)={p1.DistanceTo(p2)} (drawing unit)");
                    }
                }
                tr.Commit();
            }

            // scan PnPDatabase Fasteners table
            PnPDatabase pnpDb;
            try { pnpDb = dlm.GetPnPDatabase(); }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] GetPnPDatabase failed: {ex.Message}");
                return;
            }

            var table = pnpDb.Tables["Fasteners"];
            if (table == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] PnPTable 'Fasteners' not found.");
                return;
            }

            object rowsObj = table.Rows;
            if (rowsObj == null)
            {
                ed.WriteMessage("\n[UFLOW][DBG] table.Rows is null.");
                return;
            }

            var colNames = GetColumnNames(table);
            bool isFallback = colNames.Count > 0 && colNames[0] == "__FALLBACK_COLUMNS__";
            if (isFallback) colNames.RemoveAt(0);
            ed.WriteMessage($"\n[UFLOW][DBG] Fasteners columns ~ {colNames.Count} (fallback={isFallback})");


            string idStr = connRowId.ToString();
            string guidStr = (connGuid ?? "").Trim();

            int hitRows = 0;
            int shown = 0;

            object selectRet = InvokeSelectRaw(rowsObj, "1=1", ed, out string retTypeName);
            IEnumerable<object> rowEnum;

            if (selectRet != null)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] rows.Select(\"1=1\") returnType={retTypeName}");
                rowEnum = EnumerateUnknownCollection(selectRet);
            }
            else
            {
                ed.WriteMessage("\n[UFLOW][DBG] Fallback: enumerate table.Rows directly (no Select or failed).");
                rowEnum = EnumerateUnknownCollection(rowsObj);
            }

            ed.WriteMessage($"\n[UFLOW][DBG] rows.Select(\"1=1\") returnType={retTypeName}");

            foreach (var r in EnumerateUnknownCollection(selectRet))
            {
                int rowId = TryGetIntProp(r, "RowId");
                string className = TryGetStringProp(r, "ClassName")
                                ?? TryGetStringProp(r, "PnPClassName")
                                ?? "";

                var hits = new List<string>();

                foreach (var c in colNames)
                {
                    string v = TryGetRowValueAsString(r, c);
                    if (string.IsNullOrWhiteSpace(v)) continue;

                    if (!string.IsNullOrEmpty(guidStr) && v.IndexOf(guidStr, StringComparison.OrdinalIgnoreCase) >= 0)
                        hits.Add($"{c}='{Trim(v)}'");
                    else if (v.Trim() == idStr)
                        hits.Add($"{c}='{Trim(v)}'");
                    else if (v.IndexOf(idStr, StringComparison.OrdinalIgnoreCase) >= 0)
                        hits.Add($"{c}='{Trim(v)}'");
                }

                if (hits.Count > 0)
                {
                    hitRows++;
                    if (shown < 30)
                    {
                        ed.WriteMessage($"\n[UFLOW][DBG] HIT FastenerRowId={rowId} Class='{className}' Hits={hits.Count}");
                        foreach (var h in hits.Take(12))
                            ed.WriteMessage($"\n[UFLOW][DBG]   {h}");
                        if (hits.Count > 12)
                            ed.WriteMessage($"\n[UFLOW][DBG]   ... (more {hits.Count - 12})");
                        shown++;
                    }
                }
            }

            ed.WriteMessage($"\n[UFLOW][DBG] Scan done. HitRows={hitRows} (shown {shown})");
            ed.WriteMessage("\n[UFLOW][DBG] HitRows=0 の場合：FastenersテーブルがConnectorを直接参照していない（別テーブル経由）の可能性が高いです。");
        }

        // -------------------- Select/Enumerate helpers --------------------

        private static object InvokeSelectRaw(object rowsObj, string filter, Editor ed, out string retTypeName)
        {
            retTypeName = "(unknown)";
            try
            {
                if (rowsObj == null)
                {
                    ed.WriteMessage("\n[UFLOW][DBG] rowsObj is null");
                    return null;
                }

                var t = rowsObj.GetType();
                ed.WriteMessage($"\n[UFLOW][DBG] rowsObj type={t.FullName}");

                var mi = t.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                        .FirstOrDefault(m =>
                            m.Name == "Select" &&
                            m.GetParameters().Length == 1 &&
                            m.GetParameters()[0].ParameterType == typeof(string));

                if (mi == null)
                {
                    ed.WriteMessage("\n[UFLOW][DBG] rowsObj has no Select(string) method.");
                    // Select() (no arg) などがあるかも見る
                    var cand = t.GetMethods().Where(m => m.Name == "Select").ToList();
                    if (cand.Count > 0)
                    {
                        ed.WriteMessage($"\n[UFLOW][DBG] Select overloads found={cand.Count}");
                        foreach (var m in cand.Take(5))
                        {
                            var ps = m.GetParameters();
                            ed.WriteMessage($"\n[UFLOW][DBG]  Select overload: ({string.Join(",", ps.Select(p => p.ParameterType.Name))})");
                        }
                    }
                    return null;
                }

                object ret = mi.Invoke(rowsObj, new object[] { filter });
                if (ret == null)
                {
                    ed.WriteMessage("\n[UFLOW][DBG] Select(string) returned null.");
                    return null;
                }

                retTypeName = ret.GetType().FullName;
                return ret;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] Select invoke exception: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }


        /// <summary>
        /// IEnumerable 未実装の独自コレクションでも回せるようにする
        /// 優先順：IEnumerable → GetEnumerator() → Count + Item[int]/get_Item(int)
        /// </summary>
        private static IEnumerable<object> EnumerateUnknownCollection(object collection)
        {
            // ★yield を try/catch の外でだけ使うため、先に回収する
            var items = new List<object>();
            if (collection == null) return items;

            // 1) IEnumerable の場合
            if (collection is IEnumerable en)
            {
                foreach (var x in en) items.Add(x);
                return items;
            }

            var t = collection.GetType();

            // 2) GetEnumerator() がある場合（独自コレクション）
            try
            {
                var miEnum = t.GetMethod("GetEnumerator", Type.EmptyTypes);
                if (miEnum != null)
                {
                    var e = miEnum.Invoke(collection, null);
                    if (e != null)
                    {
                        var miMove = e.GetType().GetMethod("MoveNext", Type.EmptyTypes);
                        var piCurr = e.GetType().GetProperty("Current");
                        if (miMove != null && piCurr != null)
                        {
                            while ((bool)miMove.Invoke(e, null))
                            {
                                items.Add(piCurr.GetValue(e, null));
                            }
                            return items;
                        }
                    }
                }
            }
            catch
            {
                // ignore and fallthrough
            }

            // 3) Count + Item[int] / get_Item(int) がある場合
            int count = -1;
            try
            {
                var piCount = t.GetProperty("Count");
                if (piCount != null)
                {
                    var v = piCount.GetValue(collection, null);
                    if (v is int i) count = i;
                    else if (v != null && int.TryParse(v.ToString(), out var ii)) count = ii;
                }
            }
            catch { /* ignore */ }

            if (count > 0)
            {
                var piItem = t.GetProperty("Item", new[] { typeof(int) });
                var miGetItem = t.GetMethod("get_Item", new[] { typeof(int) });

                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        object item = null;
                        if (piItem != null) item = piItem.GetValue(collection, new object[] { i });
                        else if (miGetItem != null) item = miGetItem.Invoke(collection, new object[] { i });

                        if (item != null) items.Add(item);
                    }
                    catch
                    {
                        // ignore
                    }
                }
            }

            return items;
        }


        // -------------------- Other helpers --------------------

        private static string Trim(string s)
        {
            if (s == null) return "";
            s = s.Trim();
            return (s.Length > 120) ? s.Substring(0, 120) + "..." : s;
        }

        private static Dictionary<string, string> SafeGetAllProps(DataLinksManager dlm, int rowId)
        {
            try
            {
                var obj = dlm.GetAllProperties(rowId, true);
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (obj == null) return dict;

                foreach (var kv in obj)
                {
                    var keyProp = kv.GetType().GetProperty("Key");
                    var valProp = kv.GetType().GetProperty("Value");
                    if (keyProp == null || valProp == null) continue;

                    var k = keyProp.GetValue(kv)?.ToString();
                    var v = valProp.GetValue(kv)?.ToString();
                    if (!string.IsNullOrWhiteSpace(k)) dict[k] = v ?? "";
                }
                return dict;
            }
            catch
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static string GetFirstString(Dictionary<string, string> props, params string[] keys)
        {
            if (props == null) return "";
            foreach (var k in keys)
            {
                var hit = props.Keys.FirstOrDefault(x => x.Equals(k, StringComparison.OrdinalIgnoreCase));
                if (hit != null)
                {
                    var v = props[hit];
                    if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                }
            }
            return "";
        }

        private static List<string> GetColumnNames(PnPTable table)
        {
            var names = new List<string>();
            if (table == null) return names;

            try
            {
                var pi = table.GetType().GetProperty("Columns");
                var cols = pi?.GetValue(table, null) as IEnumerable;
                if (cols != null)
                {
                    foreach (var c in cols)
                    {
                        if (c == null) continue;
                        var n = c.GetType().GetProperty("Name")?.GetValue(c, null)?.ToString()
                             ?? c.GetType().GetProperty("ColumnName")?.GetValue(c, null)?.ToString();
                        if (!string.IsNullOrWhiteSpace(n)) names.Add(n);
                    }
                }
            }
            catch { }

            if (names.Count == 0)
            {
                // ここは「推測で固定」ではなく “探索用の最低限” の保険
                names.AddRange(new[]
                {
                    "RowGuid","ClassName","PnPClassName","JointType","PnPGuid","PnPID","OwnerId","ParentId","RelatedId","ConnectorId",
                    "FromRowId","ToRowId","Port1RowId","Port2RowId"
                });
                // ★保険であることが分かるように、先頭にマーカーを入れておく
                names.Insert(0, "__FALLBACK_COLUMNS__");
            }

            return names.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string TryGetRowValueAsString(object row, string colName)
        {
            try
            {
                if (row == null || string.IsNullOrWhiteSpace(colName)) return "";
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

        private static string TryGetStringProp(object obj, string name)
        {
            try
            {
                var pi = obj.GetType().GetProperty(name);
                var v = pi?.GetValue(obj, null);
                return v?.ToString();
            }
            catch { return null; }
        }

        private class PortInfo { public string Name; public Point3d Pos; }

        private static List<PortInfo> GetPortsByReflection(Entity ent)
        {
            var result = new List<PortInfo>();
            if (ent == null) return result;

            object portCollection = null;
            var t = ent.GetType();

            var mi0 = t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                      .FirstOrDefault(m => m.Name == "GetPorts" && m.GetParameters().Length == 0);
            if (mi0 != null)
            {
                try { portCollection = mi0.Invoke(ent, null); } catch { }
            }

            if (portCollection == null)
            {
                var piPorts = t.GetProperty("Ports", BindingFlags.Public | BindingFlags.Instance);
                if (piPorts != null)
                {
                    try { portCollection = piPorts.GetValue(ent, null); } catch { }
                }
            }

            if (portCollection is IEnumerable en)
            {
                foreach (var p in en)
                {
                    if (p == null) continue;
                    string name = p.GetType().GetProperty("Name")?.GetValue(p, null)?.ToString();
                    object posObj = p.GetType().GetProperty("Position")?.GetValue(p, null)
                                  ?? p.GetType().GetProperty("Location")?.GetValue(p, null);
                    Point3d pos = Point3d.Origin;
                    if (posObj is Point3d p3) pos = p3;
                    result.Add(new PortInfo { Name = name, Pos = pos });
                }
            }
            return result;
        }

        private static bool Eq(string a, string b)
            => string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
    }
}

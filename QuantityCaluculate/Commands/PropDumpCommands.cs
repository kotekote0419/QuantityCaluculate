using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
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
    /// <summary>
    /// デバッグ用：Entity / FastenerRow の全プロパティを列挙してCSVに出す。
    /// - OD候補（外径）と厚み候補（Fastener厚み）もログに抽出表示する。
    /// </summary>
    public class PropDumpCommands
    {
        [CommandMethod("UFLOW_DIAG_DUMP_PROPS")]
        public void DumpProps()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var dlm = PlantApplication.CurrentProject?.ProjectParts["Piping"]?.DataLinksManager;
            if (dlm == null)
            {
                ed.WriteMessage("\n[UFLOW][DIAG] DataLinksManagerが取得できません。");
                return;
            }

            // 1) Entity選択
            var peo = new PromptEntityOptions("\n[UFLOW][DIAG] プロパティを確認するエンティティ（Pipe/Part等）を選択：");
            peo.AllowNone = false;
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            string outPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"PropDump_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            );

            using (var sw = new StreamWriter(outPath, false, System.Text.Encoding.UTF8))
            {
                sw.WriteLine("Source,Id,Handle,RowId,PropName,Value");

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var oid = per.ObjectId;
                    var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                    string handle = ent?.Handle.ToString() ?? "";
                    int rowId = TryFindRowId(dlm, oid);

                    ed.WriteMessage($"\n[UFLOW][DIAG] Entity Handle={handle}, RowId={rowId}");

                    // Entity のプロパティ列挙
                    DumpAllPropsToCsv(sw, "Entity", oid.ToString(), handle, rowId, GetAllPropertiesAny(dlm, oid));
                    PrintCandidates(ed, "Entity", GetAllPropertiesAny(dlm, oid));

                    // 2) FastenerRowのサンプル
                    var fastenerRows = FastenerCollector.CollectFastenerRowIds(dlm, ed);
                    if (fastenerRows.Count == 0)
                    {
                        ed.WriteMessage("\n[UFLOW][DIAG] FastenerRow が 0 件のため、Fastenerのプロパティ列挙はスキップします。");
                    }
                    else
                    {
                        int selectedRowId = ChooseFastenerRowId(ed, fastenerRows);
                        ed.WriteMessage($"\n[UFLOW][DIAG] FastenerRow RowId={selectedRowId} をダンプします。");

                        DumpAllPropsToCsv(sw, "FastenerRow", selectedRowId.ToString(CultureInfo.InvariantCulture), "", selectedRowId,
                            GetAllPropertiesAny(dlm, selectedRowId));

                        PrintCandidates(ed, "FastenerRow", GetAllPropertiesAny(dlm, selectedRowId));
                    }

                    tr.Commit();
                }
            }

            ed.WriteMessage($"\n[UFLOW][DIAG] PropDump CSV 出力: {outPath}");
        }

        private static int ChooseFastenerRowId(Editor ed, List<int> fastenerRows)
        {
            // 既定：最初のrowId
            int defaultId = fastenerRows[0];

            var pso = new PromptStringOptions($"\n[UFLOW][DIAG] FastenerRowのRowIdを入力（Enterで {defaultId}）：")
            {
                AllowSpaces = false
            };
            var psr = ed.GetString(pso);
            if (psr.Status != PromptStatus.OK) return defaultId;

            var s = (psr.StringResult ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return defaultId;

            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) && id > 0)
                return id;

            ed.WriteMessage($"\n[UFLOW][DIAG] RowIdの解釈に失敗したため、既定 {defaultId} を使用します。");
            return defaultId;
        }

        // -------------------------
        // Property enumeration
        // -------------------------

        private static object GetAllPropertiesAny(DataLinksManager dlm, ObjectId oid)
        {
            // GetAllProperties(ObjectId,bool) or GetProperties(ObjectId,bool) を反射で叩く
            var props = InvokeProps(dlm, "GetAllProperties", new object[] { oid, true }, new[] { typeof(ObjectId), typeof(bool) });
            if (props != null) return props;

            props = InvokeProps(dlm, "GetProperties", new object[] { oid, true }, new[] { typeof(ObjectId), typeof(bool) });
            return props;
        }

        private static object GetAllPropertiesAny(DataLinksManager dlm, int rowId)
        {
            var props = InvokeProps(dlm, "GetAllProperties", new object[] { rowId, true }, new[] { typeof(int), typeof(bool) });
            if (props != null) return props;

            props = InvokeProps(dlm, "GetProperties", new object[] { rowId, true }, new[] { typeof(int), typeof(bool) });
            return props;
        }

        private static object InvokeProps(object target, string methodName, object[] args, Type[] sig)
        {
            try
            {
                var mi = target.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance, null, sig, null);
                if (mi == null) return null;
                return mi.Invoke(target, args);
            }
            catch
            {
                return null;
            }
        }

        private static void DumpAllPropsToCsv(StreamWriter sw, string source, string id, string handle, int rowId, object props)
        {
            foreach (var kv in EnumerateProps(props))
            {
                sw.WriteLine(string.Join(",",
                    CsvEsc(source),
                    CsvEsc(id),
                    CsvEsc(handle),
                    rowId > 0 ? rowId.ToString(CultureInfo.InvariantCulture) : "",
                    CsvEsc(kv.Key),
                    CsvEsc(kv.Value)
                ));
            }
        }

        /// <summary>
        /// DataLinksManager の GetAllProperties / GetProperties が返す様々な型から、key/value を列挙する
        /// </summary>
        private static IEnumerable<KeyValuePair<string, string>> EnumerateProps(object props)
        {
            if (props == null) yield break;

            if (props is IDictionary dict)
            {
                foreach (DictionaryEntry de in dict)
                {
                    var k = de.Key?.ToString() ?? "";
                    var v = de.Value?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(k))
                        yield return new KeyValuePair<string, string>(k, v);
                }
                yield break;
            }

            if (props is NameValueCollection nvc)
            {
                foreach (var k in nvc.AllKeys)
                {
                    if (k == null) continue;
                    yield return new KeyValuePair<string, string>(k, nvc[k] ?? "");
                }
                yield break;
            }

            if (props is IEnumerable e)
            {
                foreach (var item in e)
                {
                    if (item == null) continue;

                    var name = GetPropString(item, "Name")
                            ?? GetPropString(item, "Key")
                            ?? GetPropString(item, "PropertyName")
                            ?? GetPropString(item, "DisplayName");

                    if (string.IsNullOrWhiteSpace(name)) continue;

                    var value = GetPropObj(item, "Value")
                             ?? GetPropObj(item, "PropValue")
                             ?? GetPropObj(item, "PropertyValue");

                    var v = value?.ToString() ?? item.ToString();
                    yield return new KeyValuePair<string, string>(name, v ?? "");
                }
            }
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

        // -------------------------
        // Candidate extraction
        // -------------------------

        private static void PrintCandidates(Editor ed, string source, object props)
        {
            var all = EnumerateProps(props).ToList();
            if (all.Count == 0)
            {
                ed.WriteMessage($"\n[UFLOW][DIAG] {source}: properties=0");
                return;
            }

            var od = all.Where(p => LooksLikeOdName(p.Key)).Take(50).ToList();
            var thk = all.Where(p => LooksLikeThicknessName(p.Key)).Take(50).ToList();

            ed.WriteMessage($"\n[UFLOW][DIAG] {source}: properties={all.Count}");

            if (od.Count > 0)
            {
                ed.WriteMessage($"\n[UFLOW][DIAG] {source}: OD候補:");
                foreach (var p in od)
                    ed.WriteMessage($"\n    {p.Key} = {p.Value}");
            }
            else
            {
                ed.WriteMessage($"\n[UFLOW][DIAG] {source}: OD候補なし");
            }

            if (thk.Count > 0)
            {
                ed.WriteMessage($"\n[UFLOW][DIAG] {source}: 厚み候補:");
                foreach (var p in thk)
                    ed.WriteMessage($"\n    {p.Key} = {p.Value}");
            }
            else
            {
                ed.WriteMessage($"\n[UFLOW][DIAG] {source}: 厚み候補なし");
            }
        }

        private static bool LooksLikeOdName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            var n = name.Trim();

            // よくある候補
            if (n.IndexOf("Outside", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (n.IndexOf("Outer", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (n.IndexOf("OD", StringComparison.OrdinalIgnoreCase) >= 0) return true;

            // 日本語
            if (n.Contains("外径") || n.Contains("外 径")) return true;

            // 例: PipeOD, OD1
            if (n.StartsWith("OD", StringComparison.OrdinalIgnoreCase)) return true;

            return false;
        }

        private static bool LooksLikeThicknessName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            var n = name.Trim();

            if (n.IndexOf("Thickness", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (n.IndexOf("Thk", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (n.IndexOf("Thick", StringComparison.OrdinalIgnoreCase) >= 0) return true;

            if (n.Contains("厚み") || n.Contains("厚さ")) return true;

            // ガスケット系
            if (n.IndexOf("Gasket", StringComparison.OrdinalIgnoreCase) >= 0 &&
                (n.IndexOf("T", StringComparison.OrdinalIgnoreCase) >= 0 || n.IndexOf("Th", StringComparison.OrdinalIgnoreCase) >= 0))
                return true;

            return false;
        }

        // -------------------------
        // RowId helper
        // -------------------------

        private static int TryFindRowId(DataLinksManager dlm, ObjectId oid)
        {
            // FindAcPpRowId(ObjectId) があれば使用、なければ0
            try
            {
                var mi = dlm.GetType().GetMethod("FindAcPpRowId", BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(ObjectId) }, null);
                if (mi == null) return 0;

                var v = mi.Invoke(dlm, new object[] { oid });
                if (v == null) return 0;

                if (v is int i) return i;
                if (int.TryParse(v.ToString(), out int ii)) return ii;
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        private static string CsvEsc(string s)
        {
            s ??= "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

// Plant 3D
using Autodesk.ProcessPower.DataLinks;

namespace UFlow // ←あなたのプロジェクトnamespaceに合わせて
{
    public class DebugRowPropsCommands
    {
        [CommandMethod("UFLOW_DEBUG_DUMP_ROWPROPS")]
        public void UFLOW_DEBUG_DUMP_ROWPROPS()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            DataLinksManager dlm;
            try
            {
                dlm = DataLinksManager.GetManager(db);
            }
            catch (System.Exception ex) // ← CS0104回避
            {
                ed.WriteMessage($"\n[UFLOW][DBG] DLM get failed: {ex.Message}");
                return;
            }

            var pr = ed.GetInteger("\n[UFLOW][DBG] Input RowId (e.g. FastenerRowId): ");
            if (pr.Status != PromptStatus.OK) return;

            DumpRowProps(ed, dlm, pr.Value);
        }

        [CommandMethod("UFLOW_DEBUG_PICK_ROWPROPS")]
        public void UFLOW_DEBUG_PICK_ROWPROPS()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            DataLinksManager dlm;
            try
            {
                dlm = DataLinksManager.GetManager(db);
            }
            catch (System.Exception ex) // ← CS0104回避
            {
                ed.WriteMessage($"\n[UFLOW][DBG] DLM get failed: {ex.Message}");
                return;
            }

            var peo = new PromptEntityOptions("\n[UFLOW][DBG] Pick entity to dump its row properties: ");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            int rowId = -1;
            try
            {
                rowId = dlm.FindAcPpRowId(per.ObjectId);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] FindAcPpRowId failed: {ex.Message}");
                return;
            }

            ed.WriteMessage($"\n[UFLOW][DBG] Picked entity RowId={rowId}");
            if (rowId > 0) DumpRowProps(ed, dlm, rowId);
        }

        private static void DumpRowProps(Editor ed, DataLinksManager dlm, int rowId)
        {
            object propsObj = null;
            try
            {
                // ※GetAllPropertiesの戻り型は環境依存があるので object で受ける
                propsObj = dlm.GetAllProperties(rowId, true);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] GetAllProperties failed: {ex.Message}");
                return;
            }

            // まず Key-Value のペアに正規化する（ここがCS1061/CS0019対策の本体）
            var pairs = NormalizeToPairs(propsObj);

            if (pairs.Count == 0)
            {
                ed.WriteMessage($"\n[UFLOW][DBG] RowId={rowId} props: (empty)");
                return;
            }

            ed.WriteMessage($"\n[UFLOW][DBG] RowId={rowId} props count={pairs.Count}");

            foreach (var kv in pairs.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
            {
                string v = kv.Value ?? "";
                if (v.Length > 180) v = v.Substring(0, 180) + "...";
                ed.WriteMessage($"\n[UFLOW][DBG]  {kv.Key} = {v}");
            }

            // 接続/参照を示しそうなキーを抽出
            string[] hints = { "Owner", "Parent", "Related", "Link", "Ref", "From", "To", "Port", "Conn", "Joint", "End", "Start", "ObjectId", "Guid", "Tag" };

            var hintKeys = pairs
                .Select(p => p.Key)
                .Where(k => hints.Any(h => k.IndexOf(h, StringComparison.OrdinalIgnoreCase) >= 0))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(k => k, StringComparer.OrdinalIgnoreCase)
                .ToList(); // ← List化するので Count がプロパティとして使える（CS0019回避）

            if (hintKeys.Count > 0)
            {
                ed.WriteMessage("\n[UFLOW][DBG] ---- Hint keys (possible linkage) ----");
                foreach (var k in hintKeys)
                {
                    string v = pairs.First(p => string.Equals(p.Key, k, StringComparison.OrdinalIgnoreCase)).Value ?? "";
                    if (v.Length > 180) v = v.Substring(0, 180) + "...";
                    ed.WriteMessage($"\n[UFLOW][DBG]  {k} = {v}");
                }
            }
        }

        /// <summary>
        /// GetAllProperties戻りを「List(Key,Value)」へ正規化
        /// - IDictionary の場合：DictionaryEntryで列挙
        /// - IEnumerable の場合：要素の Key/Value or Name/Value を reflection で拾う
        /// - それ以外：ToString のみ（最後の保険）
        /// </summary>
        private static List<KV> NormalizeToPairs(object propsObj)
        {
            var list = new List<KV>();
            if (propsObj == null) return list;

            // 1) IDictionary
            if (propsObj is IDictionary dict)
            {
                foreach (DictionaryEntry de in dict)
                {
                    string k = (de.Key ?? "").ToString();
                    string v = de.Value?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(k))
                        list.Add(new KV(k, v));
                }
                return list;
            }

            // 2) IEnumerable（辞書じゃないコレクション）
            if (propsObj is IEnumerable en)
            {
                foreach (var item in en)
                {
                    if (item == null) continue;

                    // よくあるパターン：Key/Value か Name/Value
                    string k = TryGetString(item, "Key") ?? TryGetString(item, "Name");
                    string v = TryGetString(item, "Value");

                    // それでも取れないなら item.ToString() を使う
                    if (string.IsNullOrWhiteSpace(k))
                    {
                        var s = item.ToString() ?? "";
                        if (!string.IsNullOrWhiteSpace(s))
                            list.Add(new KV(s, ""));
                        continue;
                    }

                    list.Add(new KV(k, v ?? ""));
                }
                return list;
            }

            // 3) 最後の保険
            list.Add(new KV("(props)", propsObj.ToString() ?? ""));
            return list;
        }

        private static string TryGetString(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName);
                if (pi == null) return null;
                var v = pi.GetValue(obj, null);
                return v?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private readonly struct KV
        {
            public readonly string Key;
            public readonly string Value;
            public KV(string k, string v) { Key = k; Value = v; }
        }
    }
}

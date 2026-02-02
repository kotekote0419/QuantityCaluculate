using System;
using System.Collections;
using System.Collections.Generic;
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

namespace UFlowPlant3D.Commands
{
    public class DlmDebugCommands
    {
        [CommandMethod("UFLOW_DUMP_SELECTED_PROPS")]
        public void DumpSelectedProps()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var dlm = PlantApplication.CurrentProject?.ProjectParts["Piping"]?.DataLinksManager;
            if (dlm == null)
            {
                ed.WriteMessage("\n[UFLOW] DataLinksManagerが取得できません。");
                return;
            }

            var peo = new PromptEntityOptions("\n[UFLOW] 対象を1つ選択:");
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) return;

            var oid = per.ObjectId;

            int? rowId = TryGetRowId(dlm, oid);

            string outPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"UFLOW_PropsDump_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            );

            using var sw = new StreamWriter(outPath, false, System.Text.Encoding.UTF8);

            sw.WriteLine("Section,Key,Value,ValueType,Source");

            sw.WriteLine($"INFO,ObjectId,{oid.Handle},{oid.ObjectClass?.Name},");
            sw.WriteLine($"INFO,RowId,{(rowId.HasValue ? rowId.Value.ToString() : "")},,");

            // 1) GetAllProperties / GetProperties (RowId優先)
            DumpPropsByMethod(sw, dlm, oid, rowId, "GetAllProperties");
            DumpPropsByMethod(sw, dlm, oid, rowId, "GetProperties");

            // 2) GetPropertyValue を代表キーで試す（名前が違う可能性があるので、実在チェック目的）
            foreach (var key in new[] { "数量ID", "QuantityID", "QTY_ID", "LineTag", "ライン番号タグ", "MaterialCode", "材料コード", "ItemCode", "項目コード", "Size", "サイズ" })
            {
                object v1 = InvokeGetPropertyValue(dlm, rowId, oid, key);
                if (v1 != null)
                    sw.WriteLine($"GetPropertyValue,{Escape(key)},{Escape(v1.ToString())},{Escape(v1.GetType().FullName)},");
            }

            // 3) DLMに存在する SetProperties オーバーロード一覧を出す（“何を受け付けるか”確定する）
            foreach (var mi in dlm.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance).Where(m => m.Name == "SetProperties"))
            {
                var sig = string.Join(" | ", mi.GetParameters().Select(p => p.ParameterType.FullName));
                sw.WriteLine($"SetPropertiesSignature,{Escape(mi.Name)},{Escape(sig)},,");
            }

            ed.WriteMessage($"\n[UFLOW] 出力しました: {outPath}");
        }

        private static void DumpPropsByMethod(StreamWriter sw, DataLinksManager dlm, ObjectId oid, int? rowId, string methodName)
        {
            // RowId版 -> ObjectId版 の順で試す
            object props = null;
            string source = "";

            if (rowId.HasValue)
            {
                props = InvokeProps(dlm, methodName, new object[] { rowId.Value, true }, new[] { typeof(int), typeof(bool) });
                if (props != null) source = $"{methodName}(int,bool)";
            }

            if (props == null)
            {
                props = InvokeProps(dlm, methodName, new object[] { oid, true }, new[] { typeof(ObjectId), typeof(bool) });
                if (props != null) source = $"{methodName}(ObjectId,bool)";
            }

            if (props == null)
                return;

            foreach (var (k, v) in FlattenProps(props))
            {
                var vt = v?.GetType().FullName ?? "";
                sw.WriteLine($"{methodName},{Escape(k)},{Escape(v?.ToString() ?? "")},{Escape(vt)},{Escape(source)}");
            }
        }

        private static object InvokeProps(DataLinksManager dlm, string methodName, object[] args, Type[] sig)
        {
            try
            {
                var mi = dlm.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance, null, sig, null);
                return mi?.Invoke(dlm, args);
            }
            catch { return null; }
        }

        private static object InvokeGetPropertyValue(DataLinksManager dlm, int? rowId, ObjectId oid, string key)
        {
            // RowId版
            if (rowId.HasValue)
            {
                try
                {
                    var mi = dlm.GetType().GetMethod("GetPropertyValue", BindingFlags.Public | BindingFlags.Instance, null,
                        new[] { typeof(int), typeof(string) }, null);
                    if (mi != null) return mi.Invoke(dlm, new object[] { rowId.Value, key });
                }
                catch { }
            }

            // ObjectId版
            try
            {
                var mi = dlm.GetType().GetMethod("GetPropertyValue", BindingFlags.Public | BindingFlags.Instance, null,
                    new[] { typeof(ObjectId), typeof(string) }, null);
                if (mi != null) return mi.Invoke(dlm, new object[] { oid, key });
            }
            catch { }

            return null;
        }

        private static int? TryGetRowId(DataLinksManager dlm, ObjectId oid)
        {
            try
            {
                var mi = dlm.GetType().GetMethod("FindAcPpRowId", BindingFlags.Public | BindingFlags.Instance, null,
                    new[] { typeof(ObjectId) }, null);
                if (mi == null) return null;

                var v = mi.Invoke(dlm, new object[] { oid });
                if (v is int i && i > 0) return i;
                if (v != null && int.TryParse(v.ToString(), out var ii) && ii > 0) return ii;
            }
            catch { }
            return null;
        }

        private static IEnumerable<(string key, object value)> FlattenProps(object props)
        {
            if (props == null) yield break;

            // 1) IDictionary
            if (props is IDictionary dict)
            {
                foreach (DictionaryEntry de in dict)
                {
                    yield return (de.Key?.ToString() ?? "", de.Value);
                }
                yield break;
            }

            // 2) IEnumerable of something (DataLinksPropertyっぽいもの等)
            if (props is IEnumerable e)
            {
                foreach (var item in e)
                {
                    if (item == null) continue;

                    // a) item が Key/Value を持つ
                    var k = GetPropString(item, "Name")
                         ?? GetPropString(item, "Key")
                         ?? GetPropString(item, "PropertyName")
                         ?? GetPropString(item, "DisplayName");

                    object v = GetPropObj(item, "Value")
                            ?? GetPropObj(item, "PropValue")
                            ?? GetPropObj(item, "PropertyValue");

                    if (!string.IsNullOrEmpty(k))
                        yield return (k, v);

                    // b) それでも取れないなら ToString
                    else
                        yield return (item.GetType().FullName, item.ToString());
                }
                yield break;
            }

            // 3) それ以外
            yield return (props.GetType().FullName, props.ToString());
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

        private static string GetPropString(object obj, string name)
        {
            var v = GetPropObj(obj, name);
            return v?.ToString();
        }

        private static string Escape(string s)
        {
            s ??= "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.PnP3dObjects;
using UFlowPlant3D.Services;

namespace UFlowPlant3D.Commands
{
    public class MyCommands2
    {
        /// <summary>
        /// 動的に数量IDを付与（レンジ決め打ち無し）。
        /// 既存の数量IDを尊重し、maxId+1で増分採番。
        /// Pipe/InlineAsset に加えて Fasteners も対象。
        /// </summary>
        [CommandMethod("UFLOW_ADD_QTYID_DYNAMIC")]
        public void AddQuantityIdDynamic()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var dlm = Autodesk.ProcessPower.PlantInstance.PlantApplication.CurrentProject?.ProjectParts["Piping"]?.DataLinksManager;
            if (dlm == null)
            {
                ed.WriteMessage("\n[UFLOW] DataLinksManagerが取得できません。");
                return;
            }

            var targets = CollectTargets(db, dlm, ed);
            if (targets.Count == 0)
            {
                ed.WriteMessage("\n[UFLOW] 対象が見つかりませんでした。");
                return;
            }

            int updated = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (var oid in targets)
                {
                    if (oid.IsNull || oid.IsErased) continue;
                    var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    // キー生成
                    var key = QuantityKeyBuilder.BuildKey(dlm, oid, ent);
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    // 既存の数量ID（キー文字列）と同じならスキップ
                    var existing = QuantityKeyProp.Get(dlm, tr, oid);
                    if (string.Equals(existing, key, StringComparison.Ordinal)) continue;

                    // 数量IDプロパティがあればそこへ、無ければDWG XRecordへ作成・保存
                    QuantityKeyProp.Set(dlm, tr, oid, key);
                    updated++;
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n[UFLOW] 数量IDへキー反映 完了: {updated} 件");
        }


        /// <summary>
        /// Class4.cs の考え方を踏襲して、配管/機器情報をCSVへ出力（Fasteners含む）。
        /// </summary>
        [CommandMethod("UFLOW_EXPORT_COMPONENTS_CSV")]
        public void ExportComponentsCsv()
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

            // 保存先（ユーザに尋ねたい場合は PromptSaveFileName を追加してOK）
            string outPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"ComponentList_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            );

            var targets = CollectTargets(db, dlm, ed);

            using var tr = db.TransactionManager.StartTransaction();

            using var sw = new StreamWriter(outPath, false, System.Text.Encoding.UTF8);
            sw.WriteLine(string.Join(",",
                "Handle",
                "EntityType",
                "QuantityId",
                "LineTag",
                "MaterialCode",
                "ItemCode",
                "PartFamilyLongDesc",
                "Size",
                "InstallType",
                "Angle",
                "ND1", "ND2", "ND3",
                "SX", "SY", "SZ",
                "MX", "MY", "MZ",
                "EX", "EY", "EZ",
                "BX", "BY", "BZ",
                "B2X", "B2Y", "B2Z"
            ));

            int rows = 0;

            foreach (var oid in targets)
            {
                var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                var info = GeometryService.ExtractComponentInfo(dlm, oid, ent);

                // 付帯情報
                string qtyKey = QuantityKeyProp.Get(dlm, tr, oid);

                string lineTag = PlantProp.GetString(dlm, oid, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号");
                string matCode = PlantProp.GetString(dlm, oid, "MaterialCode", "材料コード", "MAT_CODE");
                string itemCode = PlantProp.GetString(dlm, oid, "ItemCode", "項目コード", "ITEM_CODE");
                string desc = PlantProp.GetString(dlm, oid, "PartFamilyLongDesc", "部品仕様詳細", "LONG_DESC", "ShortDescription");
                string size = PlantProp.GetString(dlm, oid, "Size", "サイズ", "NPS");
                string install = PlantProp.GetString(dlm, oid, "施工方法", "Installation", "INSTALLATION");
                string angle = PlantProp.GetString(dlm, oid, "Angle", "角度", "PathAngle");


                sw.WriteLine(string.Join(",",
                    Csv(info.HandleString),
                    Csv(info.EntityType),
                    key,
                    Csv(lineTag),
                    Csv(matCode),
                    Csv(itemCode),
                    Csv(desc),
                    Csv(size),
                    Csv(install),
                    Csv(angle),
                    CsvD(info.ND1), CsvD(info.ND2), CsvD(info.ND3),
                    CsvP(info.Start),
                    CsvP(info.Mid),
                    CsvP(info.End),
                    CsvP(info.Branch),
                    CsvP(info.Branch2)
                ));

                rows++;
            }

            tr.Commit();

            ed.WriteMessage($"\n[UFLOW] CSV出力: {rows} 行 -> {outPath}");
        }

        // -----------------------
        // Helpers
        // -----------------------

        private static List<ObjectId> CollectTargets(Database db, DataLinksManager dlm, Editor ed)
        {
            var targets = new List<ObjectId>();

            // ModelSpace: Pipe / Part（InlineAsset等）を対象
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId oid in ms)
                {
                    if (oid.IsNull || oid.IsErased) continue;

                    var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    if (ent is Pipe || ent is Part) // PipeInlineAsset は Part 継承
                        targets.Add(oid);
                }

                tr.Commit();
            }

            // Fasteners: PnPDatabase.Fasteners テーブルから追加
            try
            {
                var fastenerIds = FastenerCollector.CollectFastenerObjectIds(dlm);
                targets.AddRange(fastenerIds);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW] Fasteners取得で例外: {ex.Message}");
            }

            // 重複排除
            var uniq = new HashSet<ObjectId>(targets);
            return new List<ObjectId>(uniq);
        }

        /// <summary>
        /// 既存の数量IDを尊重して Key->ID を埋める。
        /// これにより「再実行でIDが変わる」事故を防ぐ。
        /// </summary>
        private static void BackfillFromExistingIds(Database db, DataLinksManager dlm, Editor ed, List<ObjectId> targets, QuantityIdStore store)
        {
            using var tr = db.TransactionManager.StartTransaction();

            foreach (var oid in targets)
            {
                if (oid.IsNull || oid.IsErased) continue;

                var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                string s = PlantProp.GetString(dlm, oid, "数量ID", "QuantityID", "QTY_ID");
                int existing = int.TryParse(s, out var tmp) ? tmp : -1;
                if (existing <= 0) continue;

                string key = QuantityKeyBuilder.BuildKey(dlm, oid, ent);
                if (string.IsNullOrWhiteSpace(key)) continue;

                // mapに無ければ既存IDで登録（衝突はログ）
                if (!store.Map.TryGetValue(key, out var mapped))
                {
                    store.Map[key] = existing;
                }
                else if (mapped != existing)
                {
                    ed.WriteMessage($"\n[UFLOW][WARN] Key衝突: {key} / mapped={mapped}, existing={existing}");
                    // 既存のmappedを優先（必要なら運用で変更可）
                }

                store.ObserveExistingId(existing);
            }

            tr.Commit();
        }

        private static bool TrySetQuantityId(DataLinksManager dlm, ObjectId oid, int id)
        {
            string v = id.ToString(CultureInfo.InvariantCulture);

            // RowId を取れれば優先
            int? rowId = null;
            try
            {
                var miRow = dlm.GetType().GetMethod("FindAcPpRowId", new[] { typeof(ObjectId) });
                if (miRow != null)
                {
                    var tmp = miRow.Invoke(dlm, new object[] { oid });
                    if (tmp is int i && i > 0) rowId = i;
                }
            }
            catch { }

            foreach (var prop in new[] { "数量ID", "QuantityID", "QTY_ID" })
            {
                // 1) SetProperties(int, string[], string[])
                if (rowId.HasValue && InvokeSet(dlm, rowId.Value, new[] { prop }, new[] { v })) return true;
                // 2) SetProperties(ObjectId, string[], string[])
                if (InvokeSet(dlm, oid, new[] { prop }, new[] { v })) return true;
                // 3) SetProperties(int, string[], object[])
                if (rowId.HasValue && InvokeSet(dlm, rowId.Value, new[] { prop }, new object[] { v })) return true;
                // 4) SetProperties(ObjectId, string[], object[])
                if (InvokeSet(dlm, oid, new[] { prop }, new object[] { v })) return true;
            }

            return false;
        }

        private static bool InvokeSet(DataLinksManager dlm, object firstArg, string[] names, object vals)
        {
            try
            {
                foreach (var mi in dlm.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance))
                {
                    if (mi.Name != "SetProperties") continue;
                    var ps = mi.GetParameters();
                    if (ps.Length != 3) continue;

                    if (!ps[0].ParameterType.IsInstanceOfType(firstArg)) continue;
                    if (ps[1].ParameterType != typeof(string[])) continue;
                    if (!ps[2].ParameterType.IsInstanceOfType(vals)) continue;

                    mi.Invoke(dlm, new object[] { firstArg, names, vals });
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static string Csv(string s)
        {
            s ??= "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static string CsvD(double? d)
            => d.HasValue ? d.Value.ToString("0.###", CultureInfo.InvariantCulture) : "";

        private static string CsvP(Autodesk.AutoCAD.Geometry.Point3d? p)
        {
            if (!p.HasValue) return ",,";
            return string.Join(",",
                p.Value.X.ToString("0.###", CultureInfo.InvariantCulture),
                p.Value.Y.ToString("0.###", CultureInfo.InvariantCulture),
                p.Value.Z.ToString("0.###", CultureInfo.InvariantCulture)
            );
        }
    }
}

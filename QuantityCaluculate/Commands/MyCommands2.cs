using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
        /// 動的に数量ID（キー文字列）を付与。
        /// Pipe/InlineAsset に加えて Fasteners も対象。
        /// </summary>
        [CommandMethod("UFLOW_ADD_QTYID_DYNAMIC")]
        public void AddQuantityIdDynamic()
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

                    var key = QuantityKeyBuilder.BuildKey(dlm, oid, ent);
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    var existing = QuantityKeyProp.Get(dlm, tr, oid);
                    if (string.Equals(existing, key, StringComparison.Ordinal)) continue;

                    QuantityKeyProp.Set(dlm, tr, oid, key);
                    updated++;
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n[UFLOW] 数量IDへキー反映 完了: {updated} 件");
        }

        /// <summary>
        /// 配管/機器情報をCSVへ出力（Fasteners含む）。
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

                string qtyKey = QuantityKeyProp.Get(dlm, tr, oid);
                if (string.IsNullOrWhiteSpace(qtyKey))
                {
                    qtyKey = QuantityKeyBuilder.BuildKey(dlm, oid, ent);
                }

                string lineTag = PlantProp.GetString(dlm, oid, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号");
                string matCode = PlantProp.GetString(dlm, oid, "MaterialCode", "材料コード", "MAT_CODE");
                string itemCode = PlantProp.GetString(dlm, oid, "ItemCode", "項目コード", "ITEM_CODE");
                string desc = PlantProp.GetString(dlm, oid, "PartFamilyLongDesc", "部品仕様詳細", "LONG_DESC", "ShortDescription");
                string size = PlantProp.GetString(dlm, oid, "Size", "サイズ", "NPS");
                string install = PlantProp.GetString(dlm, oid, "施工方法", "Installation", "INSTALLATION");
                string angle = PlantProp.GetString(dlm, oid, "Angle", "角度", "PathAngle");

                sw.WriteLine(string.Join(",",
                    CsvEsc(info.HandleString),
                    CsvEsc(info.EntityType),
                    CsvEsc(qtyKey),
                    CsvEsc(lineTag),
                    CsvEsc(matCode),
                    CsvEsc(itemCode),
                    CsvEsc(desc),
                    CsvEsc(size),
                    CsvEsc(install),
                    CsvEsc(angle),
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

                    if (ent is Pipe || ent is Part)
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
            catch (Exception ex)
            {
                ed.WriteMessage($"\n[UFLOW] Fasteners取得で例外: {ex.Message}");
            }

            return new List<ObjectId>(new HashSet<ObjectId>(targets));
        }

        private static string CsvEsc(string s)
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

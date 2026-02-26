using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.PnP3dObjects;
using ClosedXML.Excel;
using UFlowPlant3D.Services;

namespace UFlowPlant3D.Commands
{
    /// <summary>
    /// UFLOW: 数量ID（採番）付与 / コンポーネントCSV / 集計CSV を出力するコマンド群
    ///
    /// ※重要: 依存関係があるため、QuantityKeyProp / QuantityIdUtil / QuantityKeyBuilder 等の
    ///        “実在するメソッド名”だけを呼ぶ（勝手に関数名を作らない）。
    /// </summary>
    public class MyCommands2
    {
        // 必要な数量IDの総数 N（=最大ID）
        // i桁は N の桁数（例: 6000 -> 4桁 0001〜6000）
        private const int QTYID_START = 1;

        // 集計表（添付イメージ想定）の列名（LineTagの列）
        // ※実際の列セットはユーザー側のExcelに合わせて増減OK
        private static readonly string[] SUMMARY_COLUMNS = new[]
        {
            "数量ID","工種","単位","メーカー範囲",
            "清循環給水-07","浄水-07","防火水槽補給-07","加熱炉補給水-07","清循環戻水-07","SP水張-07",
            "不断水-07","消火-07","ドレン-07","範囲外","逆洗排水-07","blank","合 計"
        };

        // ----------------------------
        // 1) 数量ID（採番）付与
        // ----------------------------
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

            // 対象（Pipe/Part。Connector除外）は専用コレクタを使用
            var entityTargets = EntityTargetCollector.Collect(db);

            // FastenerRow（Gasket/BoltSet: in-dwg。座標取得用にインスタンスも保持）
            var fastenerInstances = FastenerCollector.CollectGasketBoltSetInstances(db, dlm, ed);
            var fastenerInstByRowId = fastenerInstances
                .GroupBy(x => x.RowId)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(x => (x.S1.HasValue ? 1 : 0) + (x.S2.HasValue ? 1 : 0)).First()
                );

            var fastenerRows = fastenerInstByRowId.Keys.ToList();
            if (fastenerRows.Count == 0)
            {
                // 念のため（現状は in-dwg 方式のため、通常ここには来ない）
                fastenerRows = FastenerCollector.CollectFastenerRowIds(db, dlm, ed);
            }

            ed.WriteMessage($"\n[UFLOW] Targets: Entity={entityTargets.Count}, FastenerRows={fastenerRows.Count}");
            int updated = 0, skipped = 0, failed = 0, newlyAllocated = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                int totalTargets = entityTargets.Count + fastenerRows.Count;

                // 運用B（欠番埋め）でも、既に 1..320 のような採番済みがある場合があるため、
                // MaxId は「今回対象数」ではなく「既存最大ID」も考慮して不足しない値にする。
                int existingMaxId = 0;

                // Entity側の既存 QuantityId
                foreach (var oid in entityTargets)
                {
                    try
                    {
                        string q = (QuantityKeyProp.GetQuantityId(dlm, tr, oid) ?? "").Trim();
                        if (int.TryParse(q, NumberStyles.Integer, CultureInfo.InvariantCulture, out int qi) && qi > existingMaxId)
                            existingMaxId = qi;
                    }
                    catch { /* ignore */ }
                }

                // FastenerRow側の既存 QuantityId
                foreach (int rowId in fastenerRows)
                {
                    try
                    {
                        string q = (QuantityKeyProp.GetRowQuantityId(dlm, rowId) ?? "").Trim();
                        if (int.TryParse(q, NumberStyles.Integer, CultureInfo.InvariantCulture, out int qi) && qi > existingMaxId)
                            existingMaxId = qi;
                    }
                    catch { /* ignore */ }
                }

                int maxIdForState = Math.Max(totalTargets, existingMaxId);
                if (maxIdForState < QTYID_START) maxIdForState = QTYID_START;

                QuantityIdUtil.State st = QuantityIdUtil.LoadState(tr, db, maxIdForState, QTYID_START);
                SetStateMaxId(st, maxIdForState);
                var usedIds = new HashSet<int>(st.Map.Values);
                int stateMaxId = GetStateMaxId(st, maxIdForState);



                // --- Entity（Pipe/Part）
                foreach (var oid in entityTargets)
                {
                    Entity ent = null;
                    try
                    {
                        ent = tr.GetObject(oid, OpenMode.ForWrite) as Entity;
                        if (ent == null) { skipped++; continue; }

                        // 集計キー（=採番の元キー）
                        string key = QuantityKeyBuilder.BuildKey(dlm, oid, ent) ?? "";
                        key = key.Trim();
                        if (string.IsNullOrWhiteSpace(key)) { skipped++; continue; }

                        bool had = st.Map.ContainsKey(key);
                        int id = GetOrCreateIdFillGaps(st, key, usedIds, QTYID_START, stateMaxId);
                        if (!had) newlyAllocated++;

                        string qtyId = QuantityIdUtil.Format(st, id);

                        // 既存値チェック（数量IDプロパティ or XRecord）
                        string existingQtyId = (QuantityKeyProp.GetQuantityId(dlm, tr, oid) ?? "").Trim();
                        if (string.Equals(existingQtyId, qtyId, StringComparison.Ordinal))
                        {
                            skipped++;
                            continue;
                        }

                        // 数量IDを書き込み（Plant側プロパティが無ければXRecordへフォールバック）
                        // キーも保存（XRecord）
                        QuantityKeyProp.SetKey(dlm, tr, oid, key);
                        // 任意: Project Setupで数量キー項目を作っている場合は書き込み
                        QuantityKeyProp.TryWritePlantQuantityKey(dlm, oid, key);

                        QuantityKeyProp.SetQuantityId(dlm, tr, oid, qtyId);

                        updated++;
                    }
                    catch (InvalidOperationException ex)
                    {
                        ed.WriteMessage($"\n[UFLOW] QtyID overflow: {ex.Message}");
                        failed++;
                        break;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n[UFLOW] Entity QtyID set failed: Handle={(ent != null ? ent.Handle.ToString() : oid.ToString())} Type={(ent != null ? ent.GetType().Name : "?")} {ex.GetType().Name} / {ex.Message}");
                        failed++;
                    }
                }

                // --- FastenerRow（rowId）
                foreach (int rowId in fastenerRows)
                {
                    try
                    {
                        string key = QuantityKeyBuilder.BuildFastenerKey(dlm, rowId) ?? "";
                        key = key.Trim();
                        if (string.IsNullOrWhiteSpace(key)) { skipped++; continue; }

                        bool had = st.Map.ContainsKey(key);
                        int id = GetOrCreateIdFillGaps(st, key, usedIds, QTYID_START, stateMaxId);
                        if (!had) newlyAllocated++;

                        string qtyId = QuantityIdUtil.Format(st, id);

                        string existingQtyId = (QuantityKeyProp.GetRowQuantityId(dlm, rowId) ?? "").Trim();
                        if (string.Equals(existingQtyId, qtyId, StringComparison.Ordinal))
                        {
                            skipped++;
                            continue;
                        }

                        // FastenerRow はXRecordフォールバックが無いので、Plant側にプロパティが無い場合は何も起きない
                        // （現状の運用では “一旦このままOK” という合意の前提）
                        // 任意: 行側に数量キー項目がある場合は書き込み
                        QuantityKeyProp.TryWriteRowQuantityKey(dlm, rowId, key);
                        QuantityKeyProp.SetRowQuantityId(dlm, rowId, qtyId);

                        updated++;
                    }
                    catch (InvalidOperationException ex)
                    {
                        ed.WriteMessage($"\n[UFLOW] QtyID overflow: {ex.Message}");
                        failed++;
                        break;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n[UFLOW] FastenerRow QtyID set failed (RowId={rowId}): {ex.GetType().Name} / {ex.Message}");
                        failed++;
                    }
                }

                SetStateNextId(st, usedIds.Count == 0 ? QTYID_START : (usedIds.Max() + 1));
                QuantityIdUtil.SaveState(tr, db, st);
                tr.Commit();
            }

            ed.WriteMessage($"\n[UFLOW] 完了: 更新 {updated} / 新規採番 {newlyAllocated} / スキップ {skipped} / 失敗 {failed}");
        }

        // ----------------------------
        // 1b) 数量IDレコード全クリア（モデルプロパティ + XRecord + State）
        // ----------------------------
        [CommandMethod("UFLOW_CLEAR_QTYID_RECORDS")]
        public void ClearQtyIdRecords()
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

            var pko = new PromptKeywordOptions("\n[UFLOW] QuantityId/QuantityKey/State を全クリアします。続行しますか？");
            pko.AllowNone = false;
            pko.Keywords.Add("Yes");
            pko.Keywords.Add("No");
            pko.Keywords.Default = "No";
            var pkr = ed.GetKeywords(pko);
            if (pkr.Status != PromptStatus.OK || !string.Equals(pkr.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
            {
                ed.WriteMessage("\n[UFLOW] Cancelled.");
                return;
            }

            var entityTargets = EntityTargetCollector.Collect(db);

            var fastenerInstances = FastenerCollector.CollectGasketBoltSetInstances(db, dlm, ed);
            var fastenerRows = fastenerInstances
                .Select(x => x.RowId)
                .Distinct()
                .ToList();
            if (fastenerRows.Count == 0)
            {
                fastenerRows = FastenerCollector.CollectFastenerRowIds(db, dlm, ed);
            }

            int clearedEnt = 0, clearedRow = 0, failed = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Entity: QuantityId/Key を空に
                foreach (var oid in entityTargets)
                {
                    Entity ent = null;
                    try
                    {
                        ent = tr.GetObject(oid, OpenMode.ForWrite) as Entity;
                        if (ent == null) continue;

                        // XRecord +（可能なら）Plant側プロパティを空文字へ
                        QuantityKeyProp.SetKey(dlm, tr, oid, "");
                        QuantityKeyProp.TryWritePlantQuantityKey(dlm, oid, "");

                        QuantityKeyProp.SetQuantityId(dlm, tr, oid, "");

                        clearedEnt++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n[UFLOW] Clear Entity failed: Handle={(ent != null ? ent.Handle.ToString() : oid.ToString())} {ex.GetType().Name} / {ex.Message}");
                        failed++;
                    }
                }

                // FastenerRow: RowのQuantityId/Key を空に
                foreach (int rowId in fastenerRows)
                {
                    try
                    {
                        QuantityKeyProp.TryWriteRowQuantityKey(dlm, rowId, "");
                        QuantityKeyProp.SetRowQuantityId(dlm, rowId, "");
                        clearedRow++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\n[UFLOW] Clear FastenerRow failed: RowId={rowId} {ex.GetType().Name} / {ex.Message}");
                        failed++;
                    }
                }

                // State: map クリア + Next/Max 初期化
                try
                {
                    var st = QuantityIdUtil.LoadState(tr, db, QTYID_START, QTYID_START);
                    if (st?.Map != null) st.Map.Clear();
                    SetStateMaxId(st, QTYID_START);
                    SetStateNextId(st, QTYID_START);
                    QuantityIdUtil.SaveState(tr, db, st);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n[UFLOW] Clear State failed: {ex.GetType().Name} / {ex.Message}");
                    failed++;
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n[UFLOW] QtyID records cleared: Entity={clearedEnt}, FastenerRows={clearedRow}, Failed={failed}");
        }

        // ----------------------------
        // QtyID allocation (Fill gaps)
        // ----------------------------
        /// <summary>
        /// 運用B: 既存IDは維持し、未採番キーに対しては 1..maxId の欠番から採番する。
        /// </summary>
        private static int GetOrCreateIdFillGaps(
            QuantityIdUtil.State st,
            string key,
            HashSet<int> usedIds,
            int startId,
            int maxId)
        {
            if (st == null) throw new ArgumentNullException(nameof(st));
            if (string.IsNullOrWhiteSpace(key)) throw new ArgumentException("key is empty", nameof(key));
            if (usedIds == null) usedIds = new HashSet<int>();

            // already assigned
            if (st.Map != null && st.Map.TryGetValue(key, out int existing) && existing > 0)
            {
                usedIds.Add(existing);
                return existing;
            }

            // allocate smallest gap within [startId, maxId]
            int chosen = 0;
            for (int i = Math.Max(1, startId); i <= Math.Max(1, maxId); i++)
            {
                if (!usedIds.Contains(i)) { chosen = i; break; }
            }

            if (chosen <= 0)
            {
                // Keep message compatible with existing overflow log pattern
                int nextGuess = Math.Max(1, maxId) + 1;
                throw new InvalidOperationException($"QuantityId exceeded MaxId={maxId}. Next={nextGuess}");
            }

            if (st.Map != null)
                st.Map[key] = chosen;

            usedIds.Add(chosen);
            // keep NextId in a reasonable state (best-effort)
            SetStateNextId(st, chosen + 1);

            return chosen;
        }

        // ----------------------------
        // 2) コンポーネント明細CSV
        // ----------------------------
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

            double unitToMeter = GetInsunitsToMetersScale(db, ed);

            string outPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"ComponentList_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            );

            var entityTargets = EntityTargetCollector.Collect(db);
            // FastenerRow（Gasket/BoltSet: in-dwg。座標取得用にインスタンスも保持）
            var fastenerInstances = FastenerCollector.CollectGasketBoltSetInstances(db, dlm, ed);
            var fastenerInstByRowId = fastenerInstances
                .GroupBy(x => x.RowId)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(x => (x.S1.HasValue ? 1 : 0) + (x.S2.HasValue ? 1 : 0)).First()
                );

            var fastenerRows = fastenerInstByRowId.Keys.ToList();
            if (fastenerRows.Count == 0)
            {
                // 念のため（現状は in-dwg 方式のため、通常ここには来ない）
                fastenerRows = FastenerCollector.CollectFastenerRowIds(db, dlm, ed);
            }

            ed.WriteMessage($"\n[UFLOW] Export: Entity={entityTargets.Count}, FastenerRows={fastenerRows.Count}");

            // --- QuantityId / LineIndex formatting (string, zero-padded) ---
            int totalTargets = entityTargets.Count + fastenerRows.Count;
            int digitsQ = Math.Max(1, totalTargets.ToString().Length);

            // --- Build LineIndex mapping by RowId (includes Gasket/BoltSet) ---
            var allRowIds = new HashSet<int>();
            foreach (var oid in entityTargets)
            {
                int rid = SafeFindRowId(dlm, oid);
                if (rid > 0) allRowIds.Add(rid);
            }
            foreach (int rid in fastenerRows)
            {
                if (rid > 0) allRowIds.Add(rid);
            }
            var bundles = FastenerCollector.CollectConnectorBundles(db, dlm, allRowIds, ed);
            var lineIndexByRowId = BuildLineIndexByRowId(dlm, allRowIds, bundles, ed);


            using var tr = db.TransactionManager.StartTransaction();

            // ND列は「存在するPort数」に合わせて動的に作るため、先に全件スキャンしてmaxNdを決める
            int maxNd = 0;
            var ndCacheByHandle = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var ndCacheFastener = new Dictionary<int, List<string>>();

            foreach (var oid in entityTargets)
            {
                if (!oid.IsValid) continue;

                if (tr.GetObject(oid, OpenMode.ForRead, false) is not Entity ent) continue;

                // 念のため：Connectorは明細/集計対象外
                if (ent.GetType().FullName == "Autodesk.ProcessPower.PnP3dObjects.Connector") continue;

                string h = ent.Handle.ToString();
                if (ndCacheByHandle.ContainsKey(h)) continue;

                var nds = GeometryService.GetNominalDiametersList(dlm, oid, ed);
                ndCacheByHandle[h] = nds;
                if (nds != null && nds.Count > maxNd) maxNd = nds.Count;
            }

            foreach (int rowId in fastenerRows)
            {
                string lineIndex = (rowId > 0 && lineIndexByRowId.TryGetValue(rowId, out var liF)) ? liF : "";
                if (ndCacheFastener.ContainsKey(rowId)) continue;

                var nds = GeometryService.GetNominalDiametersListByRowId(dlm, rowId, ed);
                ndCacheFastener[rowId] = nds;
                if (nds != null && nds.Count > maxNd) maxNd = nds.Count;
            }

            if (maxNd < 1) maxNd = 1;

            using var sw = new StreamWriter(outPath, false, System.Text.Encoding.UTF8);

            // 明細ヘッダ（ND列はmaxNdに応じて動的に作る）
            var headerCols = new List<string>
            {
                "Handle","EntityType",
                "QuantityId","QtyKey","LineNumberTag","LineIndex","MaterialCode","ItemCode","PartFamilyLongDesc",
                "Size","InstallType","Angle"
            };
            for (int i = 1; i <= maxNd; i++) headerCols.Add($"S{i}.ND");

            headerCols.Add("PipeLength_m");

            // InstallLength は「各Port(Sn) → MidPoint」距離として出力する
            for (int i = 1; i <= maxNd; i++) headerCols.Add($"InstallLength_S{i}_m");

            // 座標は Port を優先（S1..Sn）。Port座標の後に MidPoint を出力
            for (int i = 1; i <= maxNd; i++)
            {
                headerCols.Add($"S{i}.X");
                headerCols.Add($"S{i}.Y");
                headerCols.Add($"S{i}.Z");
            }
            headerCols.AddRange(new[] { "Mid.X", "Mid.Y", "Mid.Z", "RowId" });

            sw.WriteLine(string.Join(",", headerCols));
            int rowsOut = 0;

            // --- Pipe / Part
            foreach (var oid in entityTargets)
            {
                var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                if (ent == null) continue;

                // ★CS0103対策: var にせず明示型（型推論失敗で info が消える事故を防止）
                ComponentInfo info = GeometryService.ExtractComponentInfo(dlm, oid, ent);
                int rowIdEnt = SafeFindRowId(dlm, oid);
                string lineIndex = (rowIdEnt > 0 && lineIndexByRowId.TryGetValue(rowIdEnt, out var liEnt)) ? liEnt : "";
                // P3dConnector は材料集計対象外（Connector表現の混入対策）
                if (string.Equals(info.EntityType, "P3dConnector", StringComparison.OrdinalIgnoreCase))
                {
                    ed.WriteMessage($@"\n[UFLOW][DBG] Skip P3dConnector in ComponentList: Handle={ent.Handle} RowId={rowIdEnt}");
                    continue;
                }

                string qtyId = (QuantityKeyProp.GetQuantityId(dlm, tr, oid) ?? "").Trim();
                qtyId = NormalizeNumericIdString(qtyId, digitsQ);
                string qtyKey = (QuantityKeyBuilder.BuildKey(dlm, oid, ent) ?? "").Trim();

                string lineTag = PlantProp.GetString(dlm, oid, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号") ?? "";
                string matCode = PlantProp.GetString(dlm, oid, "MaterialCode", "材料コード", "MAT_CODE") ?? "";
                string itemCode = PlantProp.GetString(dlm, oid, "ItemCode", "項目コード", "ITEM_CODE") ?? "";
                string desc = PlantProp.GetString(dlm, oid, "PartFamilyLongDesc", "部品仕様詳細", "LONG_DESC", "ShortDescription") ?? "";
                string size = PlantProp.GetString(dlm, oid, "Size", "サイズ", "NPS") ?? "";
                string install = PlantProp.GetString(dlm, oid, "施工方法", "Installation", "INSTALLATION") ?? "";
                string angle = PlantProp.GetString(dlm, oid, "Angle", "角度", "PathAngle") ?? "";


                // Size整形：ComponentListに格納する際に "mm"（全角含む）を除去する
                size = NormalizeSizeRemoveMm(size);

                // TeeのSizeは「母管x分岐管」に正規化（ComponentList用）
                bool isTeeForSize = string.Equals(info.EntityType, "Tee", StringComparison.OrdinalIgnoreCase)
                                   || ent.GetType().Name.IndexOf("Tee", StringComparison.OrdinalIgnoreCase) >= 0;
                if (isTeeForSize && ndCacheByHandle.TryGetValue(info.HandleString, out var ndsForTee))
                {
                    size = NormalizeTeeSizeMainBranch(size, ndsForTee);
                }

                double pipeLenM = 0.0;
                if (ent is Pipe && info.Start.HasValue && info.End.HasValue)
                {
                    pipeLenM = info.Start.Value.DistanceTo(info.End.Value) * unitToMeter;
                }

                // Port座標（S1..Sn）と MidPoint を取得
                var ports = GeometryService.GetPorts(ent) ?? new List<Point3d>();

                // MidPoint が取れない場合は、S1/S2 から補完（最低限）
                Point3d? mid = info.Mid;
                if (!mid.HasValue)
                {
                    if (ports.Count >= 2)
                        mid = new Point3d((ports[0].X + ports[1].X) / 2.0, (ports[0].Y + ports[1].Y) / 2.0, (ports[0].Z + ports[1].Z) / 2.0);
                    else if (info.Start.HasValue && info.End.HasValue)
                        mid = new Point3d((info.Start.Value.X + info.End.Value.X) / 2.0, (info.Start.Value.Y + info.End.Value.Y) / 2.0, (info.Start.Value.Z + info.End.Value.Z) / 2.0);
                }


                // Tee special: 人孔用=True の場合は S3 の InstallLength を 0 とする
                bool teeManhole = false;
                // Avoid compile-time dependency on a specific Tee type; rely on EntityType/name.
                if (string.Equals(info.EntityType, "Tee", StringComparison.OrdinalIgnoreCase)
                    || ent.GetType().Name.IndexOf("Tee", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string manholeStr = (PlantProp.GetString(dlm, oid, "人孔用") ?? "").Trim();
                    teeManhole = bool.TryParse(manholeStr, out bool b) && b;
                }
                sw.WriteLine(string.Join(",",
<<<<<<< HEAD
                                    CsvEsc(info.HandleString),
                                    CsvEsc(info.EntityType),
                                    CsvEscQ(qtyId),
                                    CsvEsc(qtyKey),
                                    CsvEsc(lineTag), CsvEscQ(lineIndex), CsvEsc(matCode), CsvEsc(itemCode), CsvEsc(desc), CsvEsc(size), CsvEsc(install), CsvEsc(angle),
                                    string.Join(",", BuildNdCols(maxNd, ndCacheByHandle.TryGetValue(info.HandleString, out var ndsE) ? ndsE : null)),
                                    CsvLen4(pipeLenM),
                                    string.Join(",", BuildInstallLenPortCols(maxNd, ports, mid, unitToMeter, teeManhole ? new HashSet<int> { 2 } : null)),
                                    string.Join(",", BuildPortCoordCols(maxNd, ports)),
                                    CsvP(mid),
                                    rowIdEnt.ToString(CultureInfo.InvariantCulture) // RowId
                                ));
                rowsOut++;
            }

            // --- FastenerRow（Gasket/BoltSet/Buttweld）
            foreach (int rowId in fastenerRows)
            {
                string lineIndex = (rowId > 0 && lineIndexByRowId.TryGetValue(rowId, out var liF)) ? liF : "";
                string qtyId = (QuantityKeyProp.GetRowQuantityId(dlm, rowId) ?? "").Trim();
                qtyId = NormalizeNumericIdString(qtyId, digitsQ);
                string qtyKey = (QuantityKeyBuilder.BuildFastenerKey(dlm, rowId) ?? "").Trim();

                string lineTag = PlantProp.GetString(dlm, rowId, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号") ?? "";
                string matCode = PlantProp.GetString(dlm, rowId, "MaterialCode", "材料コード", "MAT_CODE") ?? "";
                string itemCode = PlantProp.GetString(dlm, rowId, "ItemCode", "項目コード", "ITEM_CODE") ?? "";
                string desc = PlantProp.GetString(dlm, rowId, "PartFamilyLongDesc", "部品仕様詳細", "LONG_DESC", "ShortDescription", "Description", "Desc") ?? "";
                string size = PlantProp.GetString(dlm, rowId, "Size", "サイズ", "NPS") ?? "";
                string install = PlantProp.GetString(dlm, rowId, "施工方法", "Installation", "INSTALLATION") ?? "";
                string angle = PlantProp.GetString(dlm, rowId, "Angle", "角度", "PathAngle") ?? "";

                // Size整形：ComponentListに格納する際に "mm"（全角含む）を除去する
                size = NormalizeSizeRemoveMm(size);

                // FastenerRowの直管長は0
                double pipeLenM = 0.0;

                // 座標：可能なら FastenerCollector で取得した S1/S2（Gasket厚み）または SourceConnectorHandle を起点に取得
                string handleStr = "";
                List<Point3d> portsF = null;
                Point3d? midF = null;
=======
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
>>>>>>> origin/master

                if (fastenerInstByRowId != null && fastenerInstByRowId.TryGetValue(rowId, out var inst) && inst != null)
                {
                    handleStr = inst.SourceConnectorHandle ?? "";

                    portsF = new List<Point3d>();
                    if (inst.S1.HasValue) portsF.Add(inst.S1.Value);
                    if (inst.S2.HasValue) portsF.Add(inst.S2.Value);
                    // BoltSet などで S1/S2 が無い場合：座標は空のまま（InstallLength算出不要）
                }

                if (portsF != null)
                {
                    if (portsF.Count >= 2)
                        midF = new Point3d((portsF[0].X + portsF[1].X) / 2.0, (portsF[0].Y + portsF[1].Y) / 2.0, (portsF[0].Z + portsF[1].Z) / 2.0);
                    else if (portsF.Count == 1)
                        midF = portsF[0];
                }

                // EntityType: ComponentListの集計キーになるため、PnPClassName（例：BoltSet / Gasket / Buttweld）を優先
                string entityType = (PlantProp.GetString(dlm, rowId, "PnPClassName", "PnpClassName", "ClassName", "PnPClass") ?? "").Trim();
                if (string.IsNullOrWhiteSpace(entityType)) entityType = "FastenerRow";

                sw.WriteLine(string.Join(",",
                    CsvEsc(handleStr),
                    CsvEsc(entityType),
                    CsvEscQ(qtyId), CsvEsc(qtyKey),
                    CsvEsc(lineTag), CsvEscQ(lineIndex), CsvEsc(matCode), CsvEsc(itemCode), CsvEsc(desc), CsvEsc(size), CsvEsc(install), CsvEsc(angle),
                    string.Join(",", BuildNdCols(maxNd, ndCacheFastener.TryGetValue(rowId, out var ndsF) ? ndsF : null)),
                    CsvLen4(pipeLenM),
                    string.Join(",", BuildInstallLenPortCols(maxNd, portsF, midF, unitToMeter, null)),
                    string.Join(",", BuildPortCoordCols(maxNd, portsF)),
                    CsvP(midF),
                    rowId.ToString(CultureInfo.InvariantCulture)
                ));
                rowsOut++;
            }

            tr.Commit();
            ed.WriteMessage($"\n[UFLOW] CSV出力: {rowsOut} 行 -> {outPath}");
        }

        // ----------------------------
        // 3) 集計CSV（数量集計表風）
        //    - 直管長：Pipeの端点距離
        //    - 敷設長：InstallLengthServiceの寄与（Valve/Flange等も含む）
        //      ※「敷設長の項目は直管長と同じ（行セットはPipeの数量IDへ寄せる）」運用のため、
        //        非Pipeの敷設寄与は “同一LineTag+径” のPipe数量IDへ寄せる（近似）。
        // ----------------------------
        [CommandMethod("UFLOW_EXPORT_SUMMARY_CSV")]
        public void ExportSummaryCsv()
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

            double unitToMeter = GetInsunitsToMetersScale(db, ed);

            string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            string outPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"ComponentSummary_{ts}.xlsx"
            );
            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"ComponentSummary_{ts}.txt"
            );

            var entityTargets = EntityTargetCollector.Collect(db);

            // --- QuantityId formatting (string, zero-padded for display) ---
            var fastenerRowsForIndex = FastenerCollector.CollectFastenerRowIds(db, dlm, ed);
            // Fastener instances (in-DWG) for install-length contribution (mainly Gasket; BoltSet may be blank)
            var fastenerInstances = FastenerCollector.CollectGasketBoltSetInstances(db, dlm, ed)
                ?? new List<FastenerCollector.FastenerInstance>();
            var fastenerInstByRowId = fastenerInstances
                .GroupBy(x => x.RowId)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(x => (x.S1.HasValue ? 1 : 0) + (x.S2.HasValue ? 1 : 0)).First()
                );

            int totalTargets = entityTargets.Count + (fastenerRowsForIndex?.Count ?? 0);
            int digitsQ = Math.Max(1, totalTargets.ToString(CultureInfo.InvariantCulture).Length);

            // Build LineIndex mapping (includes fasteners). NOTE: Summary aggregation does NOT use LineIndex.
            // (LineIndex may differ even within the same work key: MaterialCode + Size + InstallType)
            var allRowIds = new HashSet<int>();
            foreach (var oid in entityTargets)
            {
                int rid = SafeFindRowId(dlm, oid);
                if (rid > 0) allRowIds.Add(rid);
            }
            if (fastenerRowsForIndex != null)
            {
                foreach (int rid in fastenerRowsForIndex) if (rid > 0) allRowIds.Add(rid);
            }
            var bundles = FastenerCollector.CollectConnectorBundles(db, dlm, allRowIds, ed);
            var lineIndexByRowId = BuildLineIndexByRowId(dlm, allRowIds, bundles, ed);

            // Pipe index:
            //   PipeKey = (LineNumberTag, InstallType)
            // Multiple pipes can match a PipeKey. In that case, we always pick the "first" pipe by sorting
            // (MaterialCode asc, QuantityId asc, Handle asc). If you want to force the destination, change MaterialCode.
            var pipeWinnerByKey = new Dictionary<string, PipeCandidate>(StringComparer.Ordinal);
            var pipeSeenCountByKey = new Dictionary<string, int>(StringComparer.Ordinal);
            var qtyIdToAggKey = new Dictionary<string, string>(StringComparer.Ordinal);

<<<<<<< HEAD
            // LineTag columns are shown only if present in the current model scan
            var lineTags = new SortedSet<string>(StringComparer.Ordinal);
=======
        private static string CsvEsc(string s)
        {
            s ??= "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
>>>>>>> origin/master

            // Aggregations:
            //  key = qtyId|work|unit
            //  col = LineNumberTag column (or "blank")
            //  value = length[m]
            var aggPipe = new Dictionary<string, Dictionary<string, double>>(StringComparer.Ordinal);
            var aggInstall = new Dictionary<string, Dictionary<string, double>>(StringComparer.Ordinal);

            // work label -> first aggKey for install-length rows (to avoid duplicate work rows in fallback)
            var workToAggKeyInstall = new Dictionary<string, string>(StringComparer.Ordinal);

            // Part quantity aggregation (non-pipe components)
            //  key = qtyId|work|unit
            //  value = count
            var aggPartQty = new Dictionary<string, Dictionary<string, double>>(StringComparer.Ordinal);

            int partQtyAdded = 0;
            int partQtySkippedNoQtyId = 0;
            int partQtySkippedNoWork = 0;

            // Bグループ（Olet/Valve/OrificePlate）で、同一(ItemCode, InstallType)にSize違いが混在する場合を検知する
            // key = entityType|itemCode|installType  value = distinct sizes
            var bGroupSizesByBaseKey = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            var partQtyRows = new List<PartQtyRow>(1024);

            // Logging / exclusions
            var excludedByRowId = new Dictionary<int, ExcludedComponent>(EqualityComparer<int>.Default);
            bool sawMultiPipeKey = false;

            using var tr = db.TransactionManager.StartTransaction();

            // 1) Build pipe base rows (PipeLength) + pipeWinnerByKey + lineTags
            foreach (var oid in entityTargets)
            {
                if (tr.GetObject(oid, OpenMode.ForRead) is not Entity ent) continue;
                if (ent is not Pipe) continue;

                ComponentInfo info = GeometryService.ExtractComponentInfo(dlm, oid, ent);
                int rowIdEnt = SafeFindRowId(dlm, oid);

                // Skip connectors
                if (string.Equals(info.EntityType, "P3dConnector", StringComparison.OrdinalIgnoreCase))
                    continue;

                string qtyIdRaw = (QuantityKeyProp.GetQuantityId(dlm, tr, oid) ?? "").Trim();
                string qtyId = NormalizeNumericIdString(qtyIdRaw, digitsQ);
                if (string.IsNullOrWhiteSpace(qtyId)) continue;

                string lt = NormalizeLineTag(PlantProp.GetString(dlm, oid, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号"));
                string install = (PlantProp.GetString(dlm, oid, "施工方法", "Installation", "INSTALLATION") ?? "").Trim();

                // PipeKey includes ND to avoid cross-size mixing in the same LineNumberTag
                List<string> pipeNds = GeometryService.GetNominalDiametersList(dlm, oid, ed) ?? new List<string>();
                string pipeNd = (pipeNds.Count > 0 ? (pipeNds[0] ?? "") : "").Trim();
                string pipeKey = MakePipeKey(lt, install, pipeNd);

                string matCode = (PlantProp.GetString(dlm, oid, "MaterialCode", "材料コード", "MAT_CODE") ?? "").Trim();
                string size = (PlantProp.GetString(dlm, oid, "Size", "サイズ", "NPS") ?? "").Trim();
                size = NormalizeSizeRemoveMm(size);
                string work = MakeWorkLabel(matCode, size, install);

                // LineTag columns
                string col = LineTagToColumn(lt);
                if (!string.IsNullOrWhiteSpace(lt) && !IsExcludedLineTag(lt))
                    lineTags.Add(lt);

                // PipeLength (prefer Start/End; fallback to S1-S2 distance if available)
                double lenM = 0.0;
                if (info.Start.HasValue && info.End.HasValue)
                {
                    lenM = info.Start.Value.DistanceTo(info.End.Value) * unitToMeter;
                }
                else if (InstallLengthService.TryComputeS1S2DistanceFromEntity(ent, out var distModelUnits))
                {
                    lenM = distModelUnits * unitToMeter;
                }

                string aggKey = MakeAggKey(qtyId, work, "m");
                qtyIdToAggKey[qtyId] = aggKey;

                if (!workToAggKeyInstall.ContainsKey(work))
                    workToAggKeyInstall[work] = aggKey;

                AddAgg(aggPipe, aggKey, col, lenM);
                AddAgg(aggInstall, aggKey, col, lenM); // install starts with pipe length base

                // Pipe winner selection per (LineNumberTag, InstallType)
                var cand = new PipeCandidate
                {
                    PipeKey = pipeKey,
                    QtyId = qtyId,
                    MaterialCode = matCode,
                    Handle = ent.Handle.ToString()
                };

                pipeSeenCountByKey.TryGetValue(pipeKey, out int cnt);
                pipeSeenCountByKey[pipeKey] = cnt + 1;

                if (!pipeWinnerByKey.TryGetValue(pipeKey, out var winner))
                {
                    pipeWinnerByKey[pipeKey] = cand;
                }
                else
                {
                    // Mark that we saw ambiguity somewhere (log once later)
                    sawMultiPipeKey = true;

                    if (ComparePipeCandidate(cand, winner) < 0)
                        pipeWinnerByKey[pipeKey] = cand;
                }
            }

            // 2) Add install contributions from non-pipe parts (and fasteners if applicable)
            foreach (var oid in entityTargets)
            {
                if (tr.GetObject(oid, OpenMode.ForRead) is not Entity ent) continue;
                if (ent is Pipe) continue;

                ComponentInfo info = GeometryService.ExtractComponentInfo(dlm, oid, ent);
                int rowIdEnt = SafeFindRowId(dlm, oid);

                // Skip connectors
                if (string.Equals(info.EntityType, "P3dConnector", StringComparison.OrdinalIgnoreCase))
                    continue;

                string lt = NormalizeLineTag(PlantProp.GetString(dlm, oid, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号"));
                string install = (PlantProp.GetString(dlm, oid, "施工方法", "Installation", "INSTALLATION") ?? "").Trim();
                string size = (PlantProp.GetString(dlm, oid, "Size", "サイズ", "NPS") ?? "").Trim();
                size = NormalizeSizeRemoveMm(size);
                string matCode = (PlantProp.GetString(dlm, oid, "MaterialCode", "材料コード", "MAT_CODE") ?? "").Trim();
                string entityType = info.EntityType ?? ent.GetType().Name;

                // --- Part quantity (non-pipe) collection (aggregated later) ---
                // 仕様の対象 EntityType だけを収集し、Bグループは Size混在を検知する
                string qtyIdRaw = (QuantityKeyProp.GetQuantityId(dlm, tr, oid) ?? "").Trim();
                string qtyId = NormalizeNumericIdString(qtyIdRaw, digitsQ);
                string itemCode = (PlantProp.GetString(dlm, oid, "ItemCode", "項目コード", "ITEM_CODE") ?? "").Trim();
                string desc = (PlantProp.GetString(dlm, oid, "PartFamilyLongDesc", "部品仕様詳細", "LONG_DESC", "ShortDescription", "Description", "Desc") ?? "").Trim();
                string etl = (entityType ?? "").Trim().ToLowerInvariant();

                // Tee size normalization (main x branch) for Summary/PartQty
                if (etl == "tee")
                {
                    var ndsForTee = GeometryService.GetNominalDiametersList(dlm, oid, ed) ?? new List<string>();
                    size = NormalizeTeeSizeMainBranch(size, ndsForTee);
                }

                // Bグループ（Olet/Valve/OrificePlate）で、同一(ItemCode, InstallType)にSize違いが混在する場合を検知
                if ((etl == "olet" || etl == "valve" || etl == "orificeplate") && !string.IsNullOrWhiteSpace(itemCode))
                {
                    string baseKey = $"{etl}|{itemCode}|{install}";
                    if (!bGroupSizesByBaseKey.TryGetValue(baseKey, out var set))
                    {
                        set = new HashSet<string>(StringComparer.Ordinal);
                        bGroupSizesByBaseKey[baseKey] = set;
                    }
                    set.Add((size ?? "").Trim());
                }

                if (IsPartQtySupportedEntityType(etl))
                {
                    partQtyRows.Add(new PartQtyRow
                    {
                        QtyId = qtyId,
                        LineNumberTag = lt,
                        EntityTypeLower = etl,
                        PartFamilyLongDesc = desc,
                        Size = size,
                        ItemCode = itemCode,
                        InstallType = install
                    });
                }
                else
                {
                    partQtySkippedNoWork++;
                }

                // Ports / Mid
                var ports = GeometryService.GetPorts(ent) ?? new List<Point3d>();
                Point3d? mid = info.Mid;
                if (!mid.HasValue)
                {
                    if (ports.Count >= 2)
                        mid = new Point3d((ports[0].X + ports[1].X) / 2.0, (ports[0].Y + ports[1].Y) / 2.0, (ports[0].Z + ports[1].Z) / 2.0);
                    else if (info.Start.HasValue && info.End.HasValue)
                        mid = new Point3d((info.Start.Value.X + info.End.Value.X) / 2.0, (info.Start.Value.Y + info.End.Value.Y) / 2.0, (info.Start.Value.Z + info.End.Value.Z) / 2.0);
                }

                // ND list aligned to S1,S2,... order
                List<string> nds = GeometryService.GetNominalDiametersList(dlm, oid, ed) ?? new List<string>();

                // Ensure LineNumberTag column exists even if there is no Pipe on the line
                if (!string.IsNullOrWhiteSpace(lt)) lineTags.Add(lt);

                // Determine target pipe (PipeKey = LineNumberTag + InstallType)
                var reasonsBase = new HashSet<string>(StringComparer.Ordinal);
                bool hasPipeKey = true;
                if (string.IsNullOrWhiteSpace(lt))
                {
                    reasonsBase.Add("MissingPipeKey(LineNumberTag)");
                    hasPipeKey = false;
                }
                if (string.IsNullOrWhiteSpace(install))
                {
                    reasonsBase.Add("MissingPipeKey(InstallType)");
                    hasPipeKey = false;
                }
                // target pipe is resolved per port using (LineNumberTag, InstallType, ND)
                // Add per-port contributions (S{n} -> Mid) with ND required
                // Exclusion is tracked per component (rowId) with reasons aggregated across ports.

                // Tee special: 人孔用=True の場合は S3 の寄与（InstallLength_S3_m）を 0 とする（=加算しない）
                bool teeManhole = false;
                // Avoid compile-time dependency on a specific Tee type; rely on EntityType/name.
                if (string.Equals(entityType, "Tee", StringComparison.OrdinalIgnoreCase)
                    || ent.GetType().Name.IndexOf("Tee", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string manholeStr = (PlantProp.GetString(dlm, oid, "人孔用") ?? "").Trim();
                    teeManhole = bool.TryParse(manholeStr, out bool b) && b;
                }
                for (int i = 0; i < Math.Max(ports.Count, nds.Count); i++)
                {
                    if (teeManhole && i == 2) continue;
                    string nd = (i < nds.Count ? (nds[i] ?? "") : "").Trim();

                    double? lenM = null;
                    if (mid.HasValue && i < ports.Count)
                    {
                        double d = ports[i].DistanceTo(mid.Value) * unitToMeter;
                        if (d > 0) lenM = d;
                    }

                    // If there is nothing to contribute, skip silently (length==0 is skipped; not excluded)
                    if (lenM == null || lenM.Value <= 0)
                        continue;

                    var reasons = new HashSet<string>(reasonsBase, StringComparer.Ordinal);

                    // Resolve target pipe QtyId for this port (ND required)
                    string targetQtyId = "";
                    if (!hasPipeKey)
                    {
                        // Missing key already recorded in base reasons
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(nd))
                        {
                            reasons.Add("MissingND");
                        }
                        else
                        {
                            string pipeKey = MakePipeKey(lt, install, nd);
                            if (pipeWinnerByKey.TryGetValue(pipeKey, out var winner))
                                targetQtyId = (winner.QtyId ?? "").Trim();
                            else
                                reasons.Add("NoTargetPipe");
                        }
                    }
                    // Skip zero/unknown length ports (do not count, do not log exclusion)
                    if (lenM == null || lenM.Value <= 0)
                        continue;

                    bool canAdd = !string.IsNullOrWhiteSpace(targetQtyId)
                                  && !string.IsNullOrWhiteSpace(nd)
                                  && (lenM != null && lenM.Value > 0);

                    // If no target pipe is found, but MaterialCode is present, fall back to work-type aggregation:
                    //   WorkKey = (MaterialCode, ND, InstallType)
                    bool didFallback = false;
                    if (!canAdd)
                    {
                        bool noTargetPipe = reasons.Contains("NoTargetPipe");
                        if (noTargetPipe
                            && !string.IsNullOrWhiteSpace(matCode)
                            && !string.IsNullOrWhiteSpace(nd)
                            && (lenM != null && lenM.Value > 0))
                        {
                            string ndKey = NormalizeNdKeyString(nd);
                            string workFb = MakeWorkLabel(matCode, ndKey, install);
                            string aggKeyFb = workToAggKeyInstall.TryGetValue(workFb, out var existingAggKeyFb)
                                ? existingAggKeyFb
                                : MakeAggKey("0", workFb, "m");
                            string colFb = LineTagToColumn(lt);
                            AddAgg(aggInstall, aggKeyFb, colFb, lenM.Value);
                            didFallback = true;
                        }
                    }

                    if (canAdd)
                    {
                        if (!qtyIdToAggKey.TryGetValue(targetQtyId, out var aggKey))
                        {
                            // Should not happen, but keep safe: create minimal row
                            aggKey = MakeAggKey(targetQtyId, "（未解決）", "m");
                            qtyIdToAggKey[targetQtyId] = aggKey;
                        }

                        string col = LineTagToColumn(lt);
                        AddAgg(aggInstall, aggKey, col, lenM.Value);
                    }
                    else if (!didFallback)
                    {
                        // record exclusion for this component/port
                        if (rowIdEnt > 0)
                        {
                            if (!excludedByRowId.TryGetValue(rowIdEnt, out var ex))
                            {
                                ex = new ExcludedComponent
                                {
                                    EntityType = entityType,
                                    LineNumberTag = lt,
                                    Size = size,
                                    InstallType = install,
                                    RowId = rowIdEnt
                                };
                                excludedByRowId[rowIdEnt] = ex;
                            }

                            foreach (var r in reasons) ex.Reasons.Add(r);

                            // record excluded port detail only if it represents a missing add
                            // (either has length or has ND but cannot add)
                            ex.AddPort(i + 1, nd, lenM);
                        }
                    }
                }
            }


            // 2b) Add install contributions from FastenerRow (Gasket/BoltSet)
            if (fastenerRowsForIndex != null && fastenerRowsForIndex.Count > 0)
            {
                foreach (int rowId in fastenerRowsForIndex)
                {
                    if (rowId <= 0) continue;

                    string lt = NormalizeLineTag(PlantProp.GetString(dlm, rowId, "LineNumberTag", "LineTag", "ライン番号タグ", "ライン番号"));
                    if (!string.IsNullOrWhiteSpace(lt)) lineTags.Add(lt);
                    string install = (PlantProp.GetString(dlm, rowId, "施工方法", "Installation", "INSTALLATION") ?? "").Trim();
                    string size = (PlantProp.GetString(dlm, rowId, "Size", "サイズ", "NPS") ?? "").Trim();
                    size = NormalizeSizeRemoveMm(size);
                    string matCode = (PlantProp.GetString(dlm, rowId, "MaterialCode", "材料コード", "MAT_CODE") ?? "").Trim();
                    string entityType = "FastenerRow";

                    // EntityType: ComponentListの集計キーになるため、PnPClassName（例：BoltSet / Gasket / Buttweld）を優先
                    try
                    {
                        entityType = (PlantProp.GetString(dlm, rowId, "PnPClassName", "PnpClassName", "ClassName", "PnPClass") ?? "").Trim();
                        if (string.IsNullOrWhiteSpace(entityType)) entityType = "FastenerRow";
                    }
                    catch { entityType = "FastenerRow"; }

                    // --- Part quantity (FastenerRow) collection (aggregated later) ---
                    string qtyIdRaw = (QuantityKeyProp.GetRowQuantityId(dlm, rowId) ?? "").Trim();
                    string qtyId = NormalizeNumericIdString(qtyIdRaw, digitsQ);
                    string itemCode = (PlantProp.GetString(dlm, rowId, "ItemCode", "項目コード", "ITEM_CODE") ?? "").Trim();
                    string desc = (PlantProp.GetString(dlm, rowId, "PartFamilyLongDesc", "部品仕様詳細", "LONG_DESC", "ShortDescription", "Description", "Desc") ?? "").Trim();
                    string etl = (entityType ?? "").Trim().ToLowerInvariant();

                    if (IsPartQtySupportedEntityType(etl))
                    {
                        partQtyRows.Add(new PartQtyRow
                        {
                            QtyId = qtyId,
                            LineNumberTag = lt,
                            EntityTypeLower = etl,
                            PartFamilyLongDesc = desc,
                            Size = size,
                            ItemCode = itemCode,
                            InstallType = install
                        });
                    }
                    else
                    {
                        partQtySkippedNoWork++;
                    }

                    // base reasons for missing PipeKey
                    var reasonsBase = new HashSet<string>(StringComparer.Ordinal);
                    bool hasPipeKey = true;
                    if (string.IsNullOrWhiteSpace(lt))
                    {
                        reasonsBase.Add("MissingPipeKey(LineNumberTag)");
                        hasPipeKey = false;
                    }
                    if (string.IsNullOrWhiteSpace(install))
                    {
                        reasonsBase.Add("MissingPipeKey(InstallType)");
                        hasPipeKey = false;
                    }

                    // ports & midpoint from in-DWG instance (if available)
                    List<Point3d> portsF = null;
                    Point3d? midF = null;
                    if (fastenerInstByRowId != null && fastenerInstByRowId.TryGetValue(rowId, out var inst) && inst != null)
                    {
                        portsF = new List<Point3d>();
                        if (inst.S1.HasValue) portsF.Add(inst.S1.Value);
                        if (inst.S2.HasValue) portsF.Add(inst.S2.Value);
                    }

                    if (portsF != null)
                    {
                        if (portsF.Count >= 2)
                            midF = new Point3d((portsF[0].X + portsF[1].X) / 2.0, (portsF[0].Y + portsF[1].Y) / 2.0, (portsF[0].Z + portsF[1].Z) / 2.0);
                        else if (portsF.Count == 1)
                            midF = portsF[0];
                    }

                    // ND list aligned to S1,S2,... order
                    List<string> ndsF = GeometryService.GetNominalDiametersListByRowId(dlm, rowId, ed) ?? new List<string>();

                    int maxCount = Math.Max(portsF?.Count ?? 0, ndsF.Count);
                    for (int i = 0; i < maxCount; i++)
                    {
                        string nd = (i < ndsF.Count ? (ndsF[i] ?? "") : "").Trim();

                        Point3d? p = null;
                        if (portsF != null && i < portsF.Count) p = portsF[i];

                        double? lenM = null;
                        if (p.HasValue && midF.HasValue)
                            lenM = p.Value.DistanceTo(midF.Value) * unitToMeter;

                        var reasons = new HashSet<string>(reasonsBase, StringComparer.Ordinal);

                        bool canAdd = hasPipeKey;

                        if (string.IsNullOrWhiteSpace(nd))
                        {
                            reasons.Add("MissingND");
                            canAdd = false;
                        }
                        // Skip zero/unknown length ports (do not count, do not log exclusion)
                        if (lenM == null || lenM.Value <= 0)
                            continue;
                        PipeCandidate targetWinner = null;
                        bool didFallbackF = false;
                        if (canAdd)
                        {
                            string pipeKey = MakePipeKey(lt, install, nd);
                            if (!pipeWinnerByKey.TryGetValue(pipeKey, out targetWinner))
                            {
                                // No target pipe: if MaterialCode exists, fall back to work-type aggregation (MaterialCode + ND + InstallType)
                                if (!string.IsNullOrWhiteSpace(matCode) && (lenM != null && lenM.Value > 0))
                                {
                                    string ndKey = NormalizeNdKeyString(nd);
                                    string workFb = MakeWorkLabel(matCode, ndKey, install);
                                    string aggKeyFb = workToAggKeyInstall.TryGetValue(workFb, out var existingAggKeyFb)
                                ? existingAggKeyFb
                                : MakeAggKey("0", workFb, "m");
                                    string colFb = LineTagToColumn(lt);
                                    AddAgg(aggInstall, aggKeyFb, colFb, lenM.Value);
                                    didFallbackF = true;
                                }
                                else
                                {
                                    reasons.Add("NoTargetPipe");
                                    canAdd = false;
                                }
                            }
                        }

                        if (canAdd && targetWinner != null)
                        {
                            string targetQtyId = targetWinner.QtyId;
                            if (string.IsNullOrWhiteSpace(targetQtyId))
                            {
                                reasons.Add("NoTargetPipe");
                                canAdd = false;
                            }
                            else
                            {
                                if (!qtyIdToAggKey.TryGetValue(targetQtyId, out string aggKey))
                                {
                                    aggKey = MakeAggKey(targetQtyId, "（未解決）", "m");
                                    qtyIdToAggKey[targetQtyId] = aggKey;
                                }

                                string col = LineTagToColumn(lt);
                                AddAgg(aggInstall, aggKey, col, lenM.Value);
                            }
                        }

                        if (!canAdd && !didFallbackF)
                        {
                            if (!excludedByRowId.TryGetValue(rowId, out var ex))
                            {
                                ex = new ExcludedComponent
                                {
                                    EntityType = entityType,
                                    LineNumberTag = lt,
                                    Size = size,
                                    InstallType = install,
                                    RowId = rowId
                                };
                                excludedByRowId[rowId] = ex;
                            }

                            foreach (var r in reasons) ex.Reasons.Add(r);
                            ex.AddPort(i + 1, nd, lenM);
                        }
                    }
                }
            }


            tr.Commit();


            // 2c) Aggregate part quantities after full scan (Bグループ Size混在を反映)
            var bGroupMixedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in bGroupSizesByBaseKey)
            {
                if (kv.Value != null && kv.Value.Count > 1) bGroupMixedKeys.Add(kv.Key);
            }

            foreach (var r in partQtyRows)
            {
                if (r == null) continue;

                bool bMixed = false;
                if ((r.EntityTypeLower == "olet" || r.EntityTypeLower == "valve" || r.EntityTypeLower == "orificeplate") && !string.IsNullOrWhiteSpace(r.ItemCode))
                {
                    string baseKey = $"{r.EntityTypeLower}|{r.ItemCode}|{r.InstallType}";
                    bMixed = bGroupMixedKeys.Contains(baseKey);
                }

                if (!TryBuildPartQtyDisplay(
                        r.EntityTypeLower,
                        r.PartFamilyLongDesc,
                        r.Size,
                        r.ItemCode,
                        r.InstallType,
                        bMixed,
                        out int orderIndex,
                        out string workDisplay,
                        out string unitQty))
                {
                    partQtySkippedNoWork++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(r.QtyId))
                {
                    partQtySkippedNoQtyId++;
                    continue;
                }

                string workInternal = $"{orderIndex:D2}	{workDisplay}";
                string colQty = LineTagToColumn(r.LineNumberTag);
                string aggKeyQty = MakeAggKey(r.QtyId, workInternal, unitQty);
                AddAgg(aggPartQty, aggKeyQty, colQty, 1.0);
                partQtyAdded++;
            }

            // 3) Columns: LineTags present in the model + blank + total
            var ltCols = lineTags
                .Where(lt => !IsExcludedLineTag(lt))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            ltCols.Sort(StringComparer.Ordinal);

            var columns = new List<string> { "数量ID", "工種", "単位" };
            columns.AddRange(ltCols);
            columns.Add("blank");
            columns.Add("合 計");

            // 4) QtyId digit width (>=4)
            int digits = 4;
            try
            {
                int maxId = 0;
                foreach (var k in aggInstall.Keys.Concat(aggPipe.Keys).Concat(aggPartQty.Keys))
                {
                    var p = k.Split('|');
                    if (p.Length <= 0) continue;
                    if (int.TryParse(p[0], out int n)) maxId = Math.Max(maxId, n);
                }
                digits = Math.Max(4, maxId.ToString(CultureInfo.InvariantCulture).Length);
            }
            catch { /* ignore */ }

            // 5) Write XLSX (CSV is no longer generated)
            string sourceDwgFullPath = "";
            try { sourceDwgFullPath = doc?.Name ?? ""; } catch { sourceDwgFullPath = ""; }
            WriteSummaryXlsx(outPath, sourceDwgFullPath, columns, aggPipe, aggInstall, aggPartQty, digits, ed);


            // 6) Write txt log (restored)
            using (var lw = new StreamWriter(logPath, false, System.Text.Encoding.UTF8))
            {
                lw.WriteLine($"[UFLOW] ComponentSummary log {ts}");
                lw.WriteLine($"[UFLOW] XLSX: {outPath}");

                if (sawMultiPipeKey)
                {
                    lw.WriteLine("[UFLOW][WARN] 複数のPipeとKeyが一致する延長は、はじめに一致したPipeに加算。加算先を指定する場合、MaterialCodeを変更してください");
                }

                // Part quantity (non-pipe) summary
                lw.WriteLine($"[PARTQTY] added={partQtyAdded} skipped_no_qtyid={partQtySkippedNoQtyId} skipped_no_work={partQtySkippedNoWork}");

                // Detect size-mix (B group)
                var mixed = bGroupSizesByBaseKey.Where(kv => kv.Value != null && kv.Value.Count > 1).ToList();
                if (mixed.Count > 0)
                {
                    lw.WriteLine($"[PARTQTY][WARN] Bグループで同一(ItemCode, InstallType)にSize違いが混在: {mixed.Count}");
                    foreach (var kv in mixed.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
                    {
                        var p = (kv.Key ?? "").Split('|');
                        string et = (p.Length > 0) ? p[0] : "";
                        string ic = (p.Length > 1) ? p[1] : "";
                        string it = (p.Length > 2) ? p[2] : "";
                        string sizes = string.Join(",", kv.Value.Where(s => !string.IsNullOrWhiteSpace(s)).OrderBy(s => s, StringComparer.Ordinal));
                        lw.WriteLine($"[PARTQTY][WARN] {et} | {ic} | {it} | sizes={sizes}");
                    }
                }

                int excludedComponents = excludedByRowId.Count;
                lw.WriteLine($"[INSTALLLEN][EXCLUDE] excluded_components={excludedComponents}");

                // List all excluded components (1 row per component)
                foreach (var ex in excludedByRowId.Values
                             .OrderBy(x => x.LineNumberTag ?? "", StringComparer.Ordinal)
                             .ThenBy(x => x.EntityType ?? "", StringComparer.Ordinal)
                             .ThenBy(x => x.RowId))
                {
                    string reasons = string.Join(",", ex.Reasons.OrderBy(r => r, StringComparer.Ordinal));
                    lw.WriteLine($"[INSTALLLEN][EXCLUDE] {ex.EntityType}/{ex.LineNumberTag}/{ex.Size}/{ex.InstallType}/{ex.RowId}  {reasons}");
                }

                // TOP2 details: max length port, max ND port
                var topLen = FindTopByLength(excludedByRowId.Values);
                var topNd = FindTopByNd(excludedByRowId.Values);

                // avoid duplicate: if both point to same component, pick next best for ND
                if (topLen != null && topNd != null && topLen.RowId == topNd.RowId)
                {
                    topNd = FindTopByNd(excludedByRowId.Values, excludeRowId: topLen.RowId);
                }

                WriteTopDetail(lw, "TOP_LENGTH", topLen);
                WriteTopDetail(lw, "TOP_ND", topNd);
            }

            ed.WriteMessage($"\n[UFLOW] 集計XLSX出力 -> {outPath}");
            ed.WriteMessage($"\n[UFLOW] ログ出力 -> {logPath}");

        }

        // ----------------------------
        // Helpers
        // ----------------------------
        // ----------------------------
        // Summary (install length) helpers
        // ----------------------------
        private sealed class PipeCandidate
        {
            public string PipeKey;
            public string QtyId;
            public string MaterialCode;
            public string Handle;
        }


        private sealed class PartQtyRow
        {
            public string QtyId;
            public string LineNumberTag;
            public string EntityTypeLower;
            public string PartFamilyLongDesc;
            public string Size;
            public string ItemCode;
            public string InstallType;
        }

        private static string MakePipeKey(string lineNumberTag, string installType, string ndKey)
        {
            string lt = NormalizeLineTag(lineNumberTag);
            string it = (installType ?? "").Trim();
            string nd = NormalizeNdKeyString(ndKey);
            return lt + "|" + it + "|" + nd;
        }

        private static int ComparePipeCandidate(PipeCandidate a, PipeCandidate b)
        {
            // "First match" is deterministic:
            //   MaterialCode asc -> QuantityId asc -> Handle asc
            // If you want to control the destination when multiple pipes match, change MaterialCode.
            string am = (a?.MaterialCode ?? "").Trim();
            string bm = (b?.MaterialCode ?? "").Trim();

            int c = string.CompareOrdinal(am, bm);
            if (c != 0) return c;

            int ai = int.MaxValue;
            int bi = int.MaxValue;
            if (int.TryParse((a?.QtyId ?? "").Trim(), out var an)) ai = an;
            if (int.TryParse((b?.QtyId ?? "").Trim(), out var bn)) bi = bn;

            c = ai.CompareTo(bi);
            if (c != 0) return c;

            return string.CompareOrdinal((a?.Handle ?? ""), (b?.Handle ?? ""));
        }

        private sealed class ExcludedPort
        {
            public int PortNo;          // 1-based (S1..)
            public string Nd;           // raw string
            public double? LengthM;     // Sn->Mid [m], null if unknown
        }

        private sealed class ExcludedComponent
        {
            public string EntityType;
            public string LineNumberTag;
            public string Size;
            public string InstallType;
            public int RowId;

            public readonly HashSet<string> Reasons = new(StringComparer.Ordinal);
            public readonly List<ExcludedPort> Ports = new();

            public void AddPort(int portNo, string nd, double? lenM)
            {
                if (portNo <= 0) return;

                // Dedup by PortNo (keep first non-empty ND/Length if later calls add more info)
                var p = Ports.FirstOrDefault(x => x.PortNo == portNo);
                if (p == null)
                {
                    Ports.Add(new ExcludedPort { PortNo = portNo, Nd = nd ?? "", LengthM = lenM });
                    return;
                }

                if (string.IsNullOrWhiteSpace(p.Nd) && !string.IsNullOrWhiteSpace(nd)) p.Nd = nd;
                if ((p.LengthM == null || p.LengthM <= 0) && (lenM != null && lenM > 0)) p.LengthM = lenM;
            }
        }

        private static ExcludedComponent FindTopByLength(IEnumerable<ExcludedComponent> comps, int? excludeRowId = null)
        {
            ExcludedComponent best = null;
            double bestLen = -1.0;

            foreach (var c in comps ?? Array.Empty<ExcludedComponent>())
            {
                if (excludeRowId.HasValue && c != null && c.RowId == excludeRowId.Value) continue;
                if (c?.Ports == null) continue;

                foreach (var p in c.Ports)
                {
                    if (p?.LengthM == null) continue;
                    double v = p.LengthM.Value;
                    if (v > bestLen)
                    {
                        bestLen = v;
                        best = c;
                    }
                }
            }
            return best;
        }

        private static ExcludedComponent FindTopByNd(IEnumerable<ExcludedComponent> comps, int? excludeRowId = null)
        {
            ExcludedComponent best = null;
            double bestNd = double.MinValue;
            double bestLen = -1.0;

            foreach (var c in comps ?? Array.Empty<ExcludedComponent>())
            {
                if (excludeRowId.HasValue && c != null && c.RowId == excludeRowId.Value) continue;
                if (c?.Ports == null) continue;

                foreach (var p in c.Ports)
                {
                    double nd = ParseNdNumber(p?.Nd);
                    if (double.IsNaN(nd)) continue;

                    double len = (p?.LengthM != null) ? p.LengthM.Value : 0.0;

                    if (nd > bestNd || (Math.Abs(nd - bestNd) < 1e-9 && len > bestLen))
                    {
                        bestNd = nd;
                        bestLen = len;
                        best = c;
                    }
                }
            }
            return best;
        }

        private static double ParseNdNumber(string nd)
        {
            // ND string may be "50", "50A", "50.0", etc. Extract leading number.
            if (string.IsNullOrWhiteSpace(nd)) return double.NaN;
            var s = nd.Trim();

            // Take continuous [0-9.] from start
            int n = 0;
            while (n < s.Length && (char.IsDigit(s[n]) || s[n] == '.')) n++;
            if (n <= 0) return double.NaN;

            var head = s.Substring(0, n);
            return double.TryParse(head, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : double.NaN;
        }

        private static void WriteTopDetail(StreamWriter lw, string tag, ExcludedComponent ex)
        {
            if (lw == null) return;

            if (ex == null)
            {
                lw.WriteLine($"[INSTALLLEN][EXCLUDE][{tag}] (none)");
                return;
            }

            string reasons = string.Join(",", ex.Reasons.OrderBy(r => r, StringComparer.Ordinal));
            lw.WriteLine($"[INSTALLLEN][EXCLUDE][{tag}] {ex.EntityType}/{ex.LineNumberTag}/{ex.Size}/{ex.InstallType}/{ex.RowId}  {reasons}");

            foreach (var p in (ex.Ports ?? new List<ExcludedPort>()).OrderBy(x => x.PortNo))
            {
                string nd = (p?.Nd ?? "").Trim();
                string len = (p?.LengthM != null && p.LengthM.Value > 0)
                    ? p.LengthM.Value.ToString("0.0000", CultureInfo.InvariantCulture)
                    : "";
                lw.WriteLine($"  S{p.PortNo}.ND={nd}  InstallLength_S{p.PortNo}_m={len}");
            }
        }


        private static string LookupPipeQtyId(
            Dictionary<string, string> pipeIndex,
            Dictionary<string, string> pipeIndexByLt,
            string lineTag,
            string ndKey)
        {
            string lt = (lineTag ?? "").Trim();
            string nd = (ndKey ?? "").Trim();

            if (pipeIndex.TryGetValue($"{lt}|{nd}", out var q1)) return q1;
            if (pipeIndexByLt.TryGetValue(lt, out var q2)) return q2;
            return "";
        }

        private static string NdKey(double? nd)
        {
            if (!nd.HasValue) return "";
            // PlantのNDは mm相当の数値が多いので、整数化してキー化（揺れ吸収）
            return Math.Round(nd.Value, 0).ToString("0", CultureInfo.InvariantCulture);
        }

        private static double GetInsunitsToMetersScale(Database db, Editor ed)
        {
            // 参照できる UnitsValue のみ使用（環境差でメンバー欠落があり得るため）
            try
            {
                var u = db.Insunits;
                if (u == UnitsValue.Millimeters) return 0.001;
                if (u == UnitsValue.Centimeters) return 0.01;
                if (u == UnitsValue.Meters) return 1.0;
                if (u == UnitsValue.Inches) return 0.0254;
                if (u == UnitsValue.Feet) return 0.3048;

                ed.WriteMessage($"\n[UFLOW][WARN] INSUNITS={u}（未対応）。mm扱い(×0.001)で換算します。");
                return 0.001;
            }
            catch
            {
                ed.WriteMessage($"\n[UFLOW][WARN] INSUNITS取得に失敗。mm扱い(×0.001)で換算します。");
                return 0.001;
            }
        }

        private static string MakeAggKey(string qtyId, string work, string unit)
            => $"{qtyId}|{work}|{unit}";

        private static void AddAgg(Dictionary<string, Dictionary<string, double>> agg, string key, string col, double v)
        {
            if (!agg.TryGetValue(key, out var cols))
            {
                cols = new Dictionary<string, double>(StringComparer.Ordinal);
                agg[key] = cols;
            }
            if (!cols.ContainsKey(col)) cols[col] = 0.0;
            cols[col] += v;
        }

        private static bool IsExcludedLineTag(string lineTag)
        {
            // 現状、除外するLineNumberTagはなし（「メーカー範囲」も列に出す）
            return false;
        }

        private static string NormalizeLineTag(string lineTag)
        {
            string lt = (lineTag ?? "").Replace('\u3000', ' ').Trim();
            if (string.IsNullOrWhiteSpace(lt)) return "";

            // いろいろなハイフン/マイナスを統一（Plant側の揺れ対策）
            lt = lt
                .Replace('－', '-')   // fullwidth hyphen-minus
                .Replace('−', '-')   // minus sign
                .Replace('‐', '-')   // hyphen
                .Replace('‑', '-')   // non-breaking hyphen
                .Replace('–', '-')   // en dash
                .Replace('—', '-');  // em dash

            return lt.Trim();
        }


        private static string NormalizeNdKeyString(string ndKey)
        {
            string s = (ndKey ?? "").Replace('\u3000', ' ').Trim();
            if (string.IsNullOrWhiteSpace(s)) return "";

            // Prefer numeric normalization (e.g., "80", "80.0" -> "80")
            // Try invariant first, then current culture.
            double dv;
            if (double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out dv) ||
                double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.CurrentCulture, out dv))
            {
                int iv = (int)System.Math.Round(dv, 0, System.MidpointRounding.AwayFromZero);
                return iv.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            // Fallback: extract leading digits (handles cases like "80A", "80 mm")
            // If we can parse digits, normalize to integer string; otherwise return trimmed original.
            int i = 0;
            while (i < s.Length && !char.IsDigit(s[i])) i++;
            int j = i;
            while (j < s.Length && (char.IsDigit(s[j]) || s[j] == '.')) j++;
            if (j > i)
            {
                string num = s.Substring(i, j - i);
                if (double.TryParse(num, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out dv) ||
                    double.TryParse(num, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.CurrentCulture, out dv))
                {
                    int iv = (int)System.Math.Round(dv, 0, System.MidpointRounding.AwayFromZero);
                    return iv.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            return s;
        }


        private static string LineTagToColumn(string lineTag)
        {
            string lt = NormalizeLineTag(lineTag);
            if (string.IsNullOrWhiteSpace(lt)) return "blank";
            if (IsExcludedLineTag(lt)) return "blank";
            return lt;
        }

        private static string MakeWorkLabel(string materialCode, string size, string installType)
        {
            string mc = (materialCode ?? "").Trim();
            string sz = (size ?? "").Trim();
            string it = (installType ?? "").Trim();

            // Plantの「配管および機器.xlsx」に合わせ、"STW 600 架設" の形にする
            // materialCode/size/installType のどれかが欠けても崩れないように結合する
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(mc)) parts.Add(mc);
            if (!string.IsNullOrWhiteSpace(sz)) parts.Add(sz);
            if (!string.IsNullOrWhiteSpace(it)) parts.Add(it);

            return string.Join(" ", parts);
        }

        // =========================
        // Part quantity (継手、バルブ、フランジ等) helpers
        // =========================
        private static string JoinTokens(params string[] tokens)
        {
            if (tokens == null || tokens.Length == 0) return "";
            var parts = new List<string>(tokens.Length);
            foreach (var t in tokens)
            {
                var s = (t ?? "").Trim();
                if (!string.IsNullOrWhiteSpace(s)) parts.Add(s);
            }
            return string.Join(" ", parts);
        }


        private static string NormalizeSizeRemoveMm(string size)
        {
            // "16mm" -> "16", "100mmx75mm" -> "100x75"
            if (string.IsNullOrWhiteSpace(size)) return "";
            string s = size.Trim();

            // fullwidth mm "ｍｍ" and ASCII "mm" (case-insensitive)
            s = s.Replace("ｍｍ", "", StringComparison.OrdinalIgnoreCase);
            s = s.Replace("mm", "", StringComparison.OrdinalIgnoreCase);

            // remove redundant spaces around separators (optional safe)
            s = s.Replace(" ", "");
            s = s.Replace("\u3000", "");

            return s;
        }

        private static string NormalizeTeeSizeMainBranch(string fallbackSize, List<string> nds)
        {
            // Prefer ND-based normalization: "母管x分岐管"
            // Rule:
            // - if one value appears twice among first 3 ports, that is mother(run); the remaining is branch
            // - else mother=max, branch=min
            if (nds == null || nds.Count == 0) return fallbackSize ?? "";

            var vals = nds.Take(3)
                          .Select(v => (v ?? "").Trim())
                          .Where(v => !string.IsNullOrWhiteSpace(v))
                          .Select(v => ExtractLeadingNumberToken(v))
                          .Where(v => !string.IsNullOrWhiteSpace(v))
                          .ToList();

            if (vals.Count < 2) return fallbackSize ?? "";

            string mother = null;
            string branch = null;

            var groups = vals.GroupBy(v => v).OrderByDescending(g => g.Count()).ToList();
            var dup = groups.FirstOrDefault(g => g.Count() >= 2);
            if (dup != null)
            {
                mother = dup.Key;
                branch = groups.Where(g => g.Key != mother).Select(g => g.Key).FirstOrDefault() ?? mother;
            }
            else
            {
                int maxV = int.MinValue;
                int minV = int.MaxValue;
                string maxS = null, minS = null;

                foreach (var v in vals)
                {
                    if (!int.TryParse(v, out int n)) continue;
                    if (n > maxV) { maxV = n; maxS = v; }
                    if (n < minV) { minV = n; minS = v; }
                }

                mother = maxS ?? vals[0];
                branch = minS ?? vals[vals.Count - 1];
            }

            if (string.IsNullOrWhiteSpace(mother) || string.IsNullOrWhiteSpace(branch)) return fallbackSize ?? "";
            return mother + "x" + branch;
        }

        private static string ExtractLeadingNumberToken(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            // e.g. "500A" -> "500", "100mm" (already stripped) -> "100"
            var m = System.Text.RegularExpressions.Regex.Match(s, @"(\d+)");
            return m.Success ? m.Groups[1].Value : s.Trim();
        }


        private static bool TryParseButtweldSize(string size, out int n)
        {
            n = 0;
            if (string.IsNullOrWhiteSpace(size)) return false;
            var m = System.Text.RegularExpressions.Regex.Match(size, @"(\d+)");
            if (!m.Success) return false;
            return int.TryParse(m.Groups[1].Value, out n);
        }

        private static string NormalizeSizeWithA(string size)
        {
            string s = (size ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return "";

            // 100A のように A を付与（既に A/a で終わっている場合はそのまま）
            if (s.EndsWith("A", StringComparison.OrdinalIgnoreCase)) return s;
            return s + "A";
        }

        /// <summary>
        /// 部品数量セクションの対象 EntityType 判定。
        /// </summary>
        private static bool IsPartQtySupportedEntityType(string entityTypeLower)
        {
            string et = (entityTypeLower ?? "").Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(et)) return false;

            return et == "elbow"
                   || et == "reducer"
                   || et == "tee"
                   || et == "flange"
                   || et == "blindflange"
                   || et == "gasket"
                   || et == "boltset"
                   || et == "orificeplate"
                   || et == "valve"
                   || et == "olet"
                   || et == "buttweld";
        }

        /// <summary>
        /// 仕様に基づき、部品数量（Pipe以外）の「工種表示」「単位」「表示順」を組み立てる。
        /// - Aグループ: Elbow, Reducer, Tee, Flange, BlindFlange, Gasket, BoltSet => PartFamilyLongDesc + Size + InstallType
        /// - Bグループ: Olet, Valve, OrificePlate => 基本 ItemCode + InstallType。Size混在時のみ " *Size" を付与して別行に分ける
        /// - Buttweld: 鋼管継手工 + Size(A付与) + InstallType
        /// </summary>
        private static bool TryBuildPartQtyDisplay(
            string entityTypeLower,
            string partFamilyLongDesc,
            string size,
            string itemCode,
            string installType,
            bool bGroupMixed,
            out int orderIndex,
            out string displayWork,
            out string unit)
        {
            orderIndex = 99;
            displayWork = "";
            unit = "";

            string et = (entityTypeLower ?? "").Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(et)) return false;

            // 表示順（要求仕様）
            switch (et)
            {
                case "elbow": orderIndex = 1; break;
                case "reducer": orderIndex = 2; break;
                case "tee": orderIndex = 3; break;
                case "flange": orderIndex = 4; break;
                case "blindflange": orderIndex = 5; break;
                case "gasket": orderIndex = 6; break;
                case "boltset": orderIndex = 7; break;
                case "orificeplate": orderIndex = 8; break;
                case "valve": orderIndex = 9; break;
                case "olet": orderIndex = 10; break;
                case "buttweld": orderIndex = 11; break;
            }

            // --- Buttweld
            if (et == "buttweld")
            {
                if (TryParseButtweldSize(size, out int bwSize) && bwSize < 400) return false;
                string szA = NormalizeSizeWithA(size);
                displayWork = JoinTokens("鋼管継手工", szA, installType);
                unit = "口";
                return !string.IsNullOrWhiteSpace(displayWork);
            }

            // --- Aグループ
            if (et == "elbow" || et == "reducer" || et == "tee" || et == "flange" || et == "blindflange" || et == "gasket" || et == "boltset")
            {
                displayWork = JoinTokens(partFamilyLongDesc, size, installType);

                // Unit: default "個"; specific types "枚"
                if (et == "flange" || et == "blindflange" || et == "gasket") unit = "枚";
                else unit = "個";

                return !string.IsNullOrWhiteSpace(displayWork);
            }

            // --- Bグループ
            if (et == "olet" || et == "valve" || et == "orificeplate")
            {
                string baseWork = JoinTokens(itemCode, installType);
                if (string.IsNullOrWhiteSpace(baseWork)) return false;

                if (bGroupMixed)
                {
                    string sz = (size ?? "").Trim();
                    displayWork = string.IsNullOrWhiteSpace(sz) ? (baseWork + " *") : (baseWork + " *" + sz);
                }
                else
                {
                    displayWork = baseWork;
                }

                unit = (et == "orificeplate") ? "枚" : "個";
                return !string.IsNullOrWhiteSpace(displayWork);
            }

            return false;
        }

        private static int GetWorkOrderIndexFromInternal(string workInternal)
        {
            if (string.IsNullOrEmpty(workInternal)) return int.MaxValue;
            var p = workInternal.Split('\t');
            if (p.Length >= 2 && int.TryParse(p[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int n))
                return n;
            return int.MaxValue;
        }

        private static string GetWorkDisplayFromInternal(string workInternal)
        {
            if (workInternal == null) return "";
            var p = workInternal.Split('\t');
            if (p.Length >= 2) return p[1] ?? "";
            return workInternal;
        }



        private static void WriteSectionHeader(StreamWriter sw, int colCount, string title)
        {
            var row = new string[colCount];
            row[0] = "";
            row[1] = title;
            for (int i = 2; i < colCount; i++) row[i] = "";
            sw.WriteLine(string.Join(",", row.Select(CsvEsc)));
        }

        private static void WriteAggRows(
            StreamWriter sw,
            List<string> columns,
            Dictionary<string, Dictionary<string, double>> agg,
            int qtyIdDigits,
            bool blankQtyId)
        {
            // 工種名（work）で昇順（セクション内の行ソート）
            // key = qtyId|work|unit
            var keys = agg.Keys.ToList();
            keys.Sort((a, b) =>
            {
                var ap = a.Split('|');
                var bp = b.Split('|');

                string awRaw = ap.Length > 1 ? ap[1] : "";
                string bwRaw = bp.Length > 1 ? bp[1] : "";

                int ao = GetWorkOrderIndexFromInternal(awRaw);
                int bo = GetWorkOrderIndexFromInternal(bwRaw);
                if (ao != bo) return ao.CompareTo(bo);

                string aw = GetWorkDisplayFromInternal(awRaw);
                string bw = GetWorkDisplayFromInternal(bwRaw);

                int c = CompareWorkNameNatural(aw, bw);
                if (c != 0) return c;

                string au = ap.Length > 2 ? ap[2] : "";
                string bu = bp.Length > 2 ? bp[2] : "";

                c = string.CompareOrdinal(au, bu);
                if (c != 0) return c;

                int ai = 0, bi = 0;
                int.TryParse(ap.Length > 0 ? ap[0] : "0", out ai);
                int.TryParse(bp.Length > 0 ? bp[0] : "0", out bi);

                c = ai.CompareTo(bi);
                if (c != 0) return c;

                return string.CompareOrdinal(a, b);
            });

            foreach (var key in keys)
            {
                var parts = key.Split('|');
                string qtyIdRaw = parts.Length > 0 ? parts[0] : "";
                string workRaw = parts.Length > 1 ? parts[1] : "";
                string work = GetWorkDisplayFromInternal(workRaw);
                string unit = parts.Length > 2 ? parts[2] : "";

                string qtyId = qtyIdRaw;
                if (!blankQtyId)
                {
                    if (int.TryParse(qtyIdRaw, out var n))
                        qtyId = n.ToString("D" + qtyIdDigits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                }
                else
                {
                    qtyId = "";
                }

                var cols = agg[key];
                double total = 0.0;

                var row = new List<string>(columns.Count)
                {
                    qtyId, work, unit
                };

                for (int i = 3; i < columns.Count; i++)
                {
                    string colName = columns[i];
                    if (colName == "合 計")
                    {
                        row.Add(CsvF(total));
                        continue;
                    }

                    double v = cols.TryGetValue(colName, out var vv) ? vv : 0.0;
                    total += v;
                    row.Add(CsvF(v));
                }

                sw.WriteLine(string.Join(",", row.Select(CsvEsc)));
            }
        }


        private static IEnumerable<string> BuildNdCols(int maxNd, List<string> nds)
        {
            // NDは「存在するPortをすべて」出す。足りない分は空欄で埋める。
            if (maxNd < 1) yield break;

            for (int i = 0; i < maxNd; i++)
            {
                string v = (nds != null && i < nds.Count) ? (nds[i] ?? "") : "";
                yield return CsvEsc(v);
            }
        }


        private static IEnumerable<string> BuildInstallLenPortCols(int maxPorts, List<Point3d> ports, Point3d? mid, double unitToMeter, HashSet<int> forceZeroPortIndex0 = null)
        {
            // InstallLength_Sn_m = (Sn -> MidPoint) 距離[m]
            if (maxPorts < 1) yield break;

            for (int i = 0; i < maxPorts; i++)
            {
                if (forceZeroPortIndex0 != null && forceZeroPortIndex0.Contains(i))
                {
                    yield return CsvLen4(0.0);
                    continue;
                }
                double? len = null;
                if (ports != null && mid.HasValue && i < ports.Count)
                {
                    len = ports[i].DistanceTo(mid.Value) * unitToMeter;
                }
                yield return CsvLen4(len);
            }
        }

        private static IEnumerable<string> BuildPortCoordCols(int maxPorts, List<Point3d> ports)
        {
            // S1.X,S1.Y,S1.Z,S2.X,... の順で出す。足りない分は空欄。
            if (maxPorts < 1) yield break;

            for (int i = 0; i < maxPorts; i++)
            {
                if (ports != null && i < ports.Count)
                {
                    yield return ports[i].X.ToString("0.###", CultureInfo.InvariantCulture);
                    yield return ports[i].Y.ToString("0.###", CultureInfo.InvariantCulture);
                    yield return ports[i].Z.ToString("0.###", CultureInfo.InvariantCulture);
                }
                else
                {
                    yield return "";
                    yield return "";
                    yield return "";
                }
            }
        }


        // =========================
        // LineIndex helpers
        // =========================
        private sealed class UnionFind
        {
            private readonly Dictionary<int, int> _parent = new();
            private readonly Dictionary<int, int> _rank = new();

            public UnionFind(IEnumerable<int> items)
            {
                foreach (var x in items)
                {
                    _parent[x] = x;
                    _rank[x] = 0;
                }
            }

            public int Find(int x)
            {
                if (!_parent.TryGetValue(x, out int p))
                {
                    _parent[x] = x;
                    _rank[x] = 0;
                    return x;
                }
                if (p != x) _parent[x] = Find(p);
                return _parent[x];
            }

            public void Union(int a, int b)
            {
                int ra = Find(a);
                int rb = Find(b);
                if (ra == rb) return;

                int rka = _rank[ra];
                int rkb = _rank[rb];
                if (rka < rkb) _parent[ra] = rb;
                else if (rka > rkb) _parent[rb] = ra;
                else
                {
                    _parent[rb] = ra;
                    _rank[ra] = rka + 1;
                }
            }
        }

        private static Dictionary<int, string> BuildLineIndexByRowId(DataLinksManager dlm, HashSet<int> allRowIds, List<FastenerCollector.ConnectorBundle> bundles, Editor ed)
        {
            var map = new Dictionary<int, string>();
            if (allRowIds == null || allRowIds.Count == 0) return map;

            // Build a connectivity graph primarily from P3dPartConnection (Part1<->Part2),
            // and additionally connect rowIds that appear in the same Connector bundle (for fasteners).
            //
            // NOTE:
            //  - This is NOT coordinate-based.
            //  - Neighbor rowIds that are NOT exported may still be pulled in as intermediate nodes,
            //    so we union over a "universe" superset, then assign indices back only for allRowIds.

            var universe = new HashSet<int>(allRowIds);
            if (bundles != null)
            {
                foreach (var b in bundles)
                {
                    if (b?.RowIds == null) continue;
                    foreach (int rid in b.RowIds)
                    {
                        if (rid > 0) universe.Add(rid);
                    }
                }
            }

            // BFS expand via P3dPartConnection so parts that connect through non-exported intermediates
            // still end up in the same connected component.
            var edges = new List<(int A, int B)>();
            var q = new Queue<int>(universe);
            var visited = new HashSet<int>();

            // Hard safety cap to avoid runaway on corrupted data
            const int MAX_VISIT = 200000;

            while (q.Count > 0 && visited.Count < MAX_VISIT)
            {
                int rid = q.Dequeue();
                if (rid <= 0) continue;
                if (!visited.Add(rid)) continue;

                foreach (int nb in GetRelatedRowIdsSafe(dlm, "P3dPartConnection", "Part1", rid, "Part2", ed))
                {
                    if (nb <= 0) continue;
                    edges.Add((rid, nb));
                    if (universe.Add(nb)) q.Enqueue(nb);
                }
                foreach (int nb in GetRelatedRowIdsSafe(dlm, "P3dPartConnection", "Part2", rid, "Part1", ed))
                {
                    if (nb <= 0) continue;
                    edges.Add((rid, nb));
                    if (universe.Add(nb)) q.Enqueue(nb);
                }
            }

            var uf = new UnionFind(universe);

            // Union Part connections
            foreach (var e in edges)
            {
                uf.Union(e.A, e.B);
            }

            // Union within each connector bundle (mostly connects fasteners to nearby parts/connectors)
            if (bundles != null)
            {
                foreach (var b in bundles)
                {
                    if (b?.RowIds == null) continue;
                    int first = 0;
                    foreach (int rid in b.RowIds)
                    {
                        if (rid <= 0) continue;
                        if (first == 0) { first = rid; continue; }
                        uf.Union(first, rid);
                    }
                }
            }

            // group by root (ONLY for exported rowIds)
            var groups = new Dictionary<int, List<int>>();
            foreach (int rid in allRowIds)
            {
                if (rid <= 0) continue;
                int root = uf.Find(rid);
                if (!groups.TryGetValue(root, out var list))
                {
                    list = new List<int>();
                    groups[root] = list;
                }
                list.Add(rid);
            }

            // stable ordering: representative = min RowId in the component (among exported ids)
            var roots = groups
                .Select(kvp => new { Root = kvp.Key, Rep = kvp.Value.Min() })
                .OrderBy(x => x.Rep)
                .ToList();

            int n = roots.Count;
            int digits = Math.Max(1, n.ToString().Length);

            for (int i = 0; i < roots.Count; i++)
            {
                string s = (i + 1).ToString($"D{digits}");
                var comp = groups[roots[i].Root];
                foreach (int rid in comp)
                {
                    map[rid] = s;
                }
            }

            return map;
        }

        private static List<int> GetRelatedRowIdsSafe(DataLinksManager dlm, string relType, string roleFrom, int rowId, string roleTo, Editor ed)
        {
            var r = new List<int>();
            if (dlm == null || rowId <= 0) return r;

            try
            {
                // DataLinksManager.GetRelatedRowIds(string relName, string role1, int rowId, string role2)
                var ids = dlm.GetRelatedRowIds(relType, roleFrom, rowId, roleTo);
                if (ids != null)
                {
                    foreach (int x in ids)
                    {
                        if (x > 0) r.Add(x);
                    }
                }
            }
            catch
            {
                // ignore: RelationshipTypeDoesNotExist / RoleDoesNotExist etc.
                // This helper is called with known-good rel/roles, so failures typically mean "not applicable to this rowId".
            }

            return r;
        }

        private static string NormalizeNumericIdString(string s, int digits)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim();
            // if already has non-digit characters, leave as is
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] < '0' || s[i] > '9') return s;
            }
            if (digits <= 0) return s;
            if (s.Length >= digits) return s;
            return s.PadLeft(digits, '0');
        }


        private static void WriteSummaryXlsx(
            string outPathXlsx,
            string sourceDwgFullPath,
            List<string> columns,
            Dictionary<string, Dictionary<string, double>> aggPipe,
            Dictionary<string, Dictionary<string, double>> aggInstall,
            Dictionary<string, Dictionary<string, double>> aggPartQty,
            int qtyIdDigits,
            Editor ed)
        {
            // NOTE:
            // - "数量集計表.xlsx" の塗りつぶし（セル色）は無視する仕様。
            // - 書式/列幅/行高/ページレイアウトは、添付の数量集計表.xlsx を参照して固定値で設定。

            using var wb = new XLWorkbook();

            var ws = wb.Worksheets.Add("集計表");

            // --- Column widths (A..Q) ---
            // A:7, B:62, C:5.5, D..Q:13 (based on attached 数量集計表.xlsx)
            ws.Column(1).Width = 7.0;
            ws.Column(2).Width = 62.0;
            ws.Column(3).Width = 5.5;
            for (int c = 4; c <= Math.Max(4, columns.Count); c++)
            {
                ws.Column(c).Width = 13.0;
            }

            // --- Default row height ---
            ws.Rows().Height = 14.25;

            // --- Page layout ---
            try
            {
                ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
                ws.PageSetup.PageOrientation = XLPageOrientation.Portrait;
                // Margins (in inches) based on attached: L/R=0.4, T/B=0.6, Header/Footer=0.5
                ws.PageSetup.Margins.Left = 0.4;
                ws.PageSetup.Margins.Right = 0.4;
                ws.PageSetup.Margins.Top = 0.6;
                ws.PageSetup.Margins.Bottom = 0.6;
                ws.PageSetup.Margins.Header = 0.8 / 2.54;// 0.8cm
                ws.PageSetup.Margins.Footer = 0.8 / 2.54;// 0.8cm

                ws.PageSetup.SetRowsToRepeatAtTop(1, 1);
                try { ws.SheetView.FreezeRows(1); } catch { } // Freeze top row

                // Fit: 1 page wide, height auto (0)
                ws.PageSetup.FitToPages(1, 0);
                // Header / Footer (match template)
                ws.PageSetup.Header.Center.AddText("数 量 集 計 表");
                ws.PageSetup.Footer.Center.AddText("&P / &Nページ");
            }
            catch (System.Exception ex)
            {
                ed?.WriteMessage("\\n[UFLOW][DBG] PageSetup skipped: " + ex.Message);
            }

            // --- Styles ---
            var styleBase = wb.Style;
            styleBase.Font.FontName = "Meiryo UI";
            styleBase.Font.FontSize = 9;
            styleBase.Font.Bold = false;
            styleBase.Alignment.Horizontal = XLAlignmentHorizontalValues.General;

            var styleHeader = wb.Style;
            styleHeader.Font.FontName = "Meiryo UI";
            styleHeader.Font.FontSize = 9;
            styleHeader.Font.Bold = true;
            styleHeader.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            styleHeader.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            var styleUnit = wb.Style;
            styleUnit.Font.FontName = "Meiryo UI";
            styleUnit.Font.FontSize = 9;
            styleUnit.Font.Bold = false;
            styleUnit.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            styleUnit.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            var styleNum = wb.Style;
            styleNum.Font.FontName = "Meiryo UI";
            styleNum.Font.FontSize = 9;
            styleNum.Font.Bold = false;
            styleNum.Alignment.Horizontal = XLAlignmentHorizontalValues.General;
            styleNum.NumberFormat.Format = "0.0";

            var styleBorder = wb.Style;
            styleBorder.Border.LeftBorder = XLBorderStyleValues.Thin;
            styleBorder.Border.RightBorder = XLBorderStyleValues.Thin;
            styleBorder.Border.TopBorder = XLBorderStyleValues.Thin;
            styleBorder.Border.BottomBorder = XLBorderStyleValues.Thin;

            // --- Header row ---
            int row = 1;
            for (int i = 0; i < columns.Count; i++)
            {
                var cell = ws.Cell(row, i + 1);
                cell.Value = columns[i];
                cell.Style = styleHeader;
                cell.Style.Border = styleBorder.Border;
            }

            // Helper: write a section header row
            void WriteSectionHeaderXlsx(string title)
            {
                row++;
                for (int c = 1; c <= columns.Count; c++)
                {
                    var cell = ws.Cell(row, c);
                    cell.Style = styleBase;
                    cell.Style.Border = styleBorder.Border;
                }
                ws.Cell(row, 2).Value = title;
            }

            // Helper: write aggregated rows; insertBlankOnOrderChange is used for part-qty entity type gaps
            void WriteAggRowsXlsx(Dictionary<string, Dictionary<string, double>> agg, bool blankQtyId, bool insertBlankOnOrderChange)
            {
                var keys = agg.Keys.ToList();
                keys.Sort((a, b) =>
                {
                    var ap = a.Split('|');
                    var bp = b.Split('|');

                    string awRaw = ap.Length > 1 ? ap[1] : "";
                    string bwRaw = bp.Length > 1 ? bp[1] : "";

                    int ao = GetWorkOrderIndexFromInternal(awRaw);
                    int bo = GetWorkOrderIndexFromInternal(bwRaw);
                    if (ao != bo) return ao.CompareTo(bo);

                    string aw = GetWorkDisplayFromInternal(awRaw);
                    string bw = GetWorkDisplayFromInternal(bwRaw);

                    int c = CompareWorkNameNatural(aw, bw);
                    if (c != 0) return c;

                    string au = ap.Length > 2 ? ap[2] : "";
                    string bu = bp.Length > 2 ? bp[2] : "";

                    c = string.CompareOrdinal(au, bu);
                    if (c != 0) return c;

                    int ai = 0, bi = 0;
                    int.TryParse(ap.Length > 0 ? ap[0] : "0", out ai);
                    int.TryParse(bp.Length > 0 ? bp[0] : "0", out bi);

                    c = ai.CompareTo(bi);
                    if (c != 0) return c;

                    return string.CompareOrdinal(a, b);
                });

                int prevOrder = int.MinValue;
                bool first = true;

                foreach (var key in keys)
                {
                    var parts = key.Split('|');
                    string qtyIdRaw = parts.Length > 0 ? parts[0] : "";
                    string workRaw = parts.Length > 1 ? parts[1] : "";
                    string work = GetWorkDisplayFromInternal(workRaw);
                    string unit = parts.Length > 2 ? parts[2] : "";

                    int curOrder = GetWorkOrderIndexFromInternal(workRaw);

                    if (insertBlankOnOrderChange && !first && curOrder != prevOrder)
                    {
                        // 1行空ける（部品数量のEntityType間）
                        row++;
                    }

                    prevOrder = curOrder;
                    first = false;

                    row++;

                    // QtyId
                    var cQty = ws.Cell(row, 1);
                    if (!blankQtyId && int.TryParse(qtyIdRaw, out var n))
                        cQty.Value = n.ToString("D" + qtyIdDigits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                    else
                        cQty.Value = "";

                    cQty.Style = styleUnit;
                    cQty.Style.Border = styleBorder.Border;

                    // Work
                    var cWork = ws.Cell(row, 2);
                    cWork.Value = work;
                    cWork.Style = styleBase;
                    cWork.Style.Border = styleBorder.Border;

                    // Unit
                    var cUnit = ws.Cell(row, 3);
                    cUnit.Value = unit;
                    cUnit.Style = styleUnit;
                    cUnit.Style.Border = styleBorder.Border;

                    var cols = agg[key];
                    double total = 0.0;

                    for (int i = 3; i < columns.Count; i++)
                    {
                        string colName = columns[i];
                        var cell = ws.Cell(row, i + 1);
                        cell.Style = styleNum;
                        cell.Style.Border = styleBorder.Border;

                        if (colName == "合 計")
                        {
                            cell.Value = total;
                            continue;
                        }

                        double v = cols.TryGetValue(colName, out var vv) ? vv : 0.0;
                        total += v;
                        cell.Value = v;
                    }
                }
            }

            // --- Sections (with 1 blank row between sections) ---
            WriteSectionHeaderXlsx("*** 直管長 ***");
            WriteAggRowsXlsx(aggPipe, blankQtyId: false, insertBlankOnOrderChange: false);

            row++; // 1行空ける（セクション間）

            WriteSectionHeaderXlsx("*** 敷設長（直管+異形管） ***");
            WriteAggRowsXlsx(aggInstall, blankQtyId: true, insertBlankOnOrderChange: false);

            row++; // 1行空ける（セクション間）

            WriteSectionHeaderXlsx("*** 部品数量（継手、バルブ、フランジ等） ***");
            WriteAggRowsXlsx(aggPartQty, blankQtyId: false, insertBlankOnOrderChange: true);

            // Source info row (1st column of the next row)
            int infoRow = row + 1;
            try
            {
                string ts = DateTime.Now.ToString("yyyy/MM/dd HH:mm", CultureInfo.InvariantCulture);
                string src = sourceDwgFullPath ?? "";
                var cInfo = ws.Cell(infoRow, 1);
                cInfo.Value = src + " - " + ts;
                cInfo.Style = styleBase;
                cInfo.Style.Font.FontSize = 8;
                cInfo.Style.Font.FontColor = XLColor.Gray;
                cInfo.Style.Border.LeftBorder = XLBorderStyleValues.None;
                cInfo.Style.Border.RightBorder = XLBorderStyleValues.None;
                cInfo.Style.Border.TopBorder = XLBorderStyleValues.None;
                cInfo.Style.Border.BottomBorder = XLBorderStyleValues.None;
            }
            catch { /* ignore */ }

            // Ensure borders for the table range only (exclude infoRow which is outside table)
            try
            {
                int tableLastRow = row;
                if (tableLastRow >= 1 && columns.Count >= 1)
                {
                    var tableRange = ws.Range(1, 1, tableLastRow, columns.Count);
                    tableRange.Style.Font.FontName = "Meiryo UI";
                    tableRange.Style.Font.FontSize = 9;
                    tableRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    tableRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    tableRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    tableRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    // Enforce alignments (test results showed some cells reverting to General)
                    try { ws.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; } catch { }
                    try { ws.Column(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; } catch { }
                    try { ws.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; } catch { }

                }
            }
            catch { /* ignore */ }

            wb.SaveAs(outPathXlsx);
        }

        private static string CsvEscQ(string s)
        {
            s ??= "";
            // Always quote to keep Excel from auto-converting leading zeros.
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }

        private static string CsvEsc(string s)
        {
            s ??= "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static int SafeFindRowId(DataLinksManager dlm, ObjectId oid)
        {
            try { return dlm.FindAcPpRowId(oid); }
            catch { return -1; }
        }



        private static bool TryGetObjectIdFromHandleString(Database db, string handleStr, out ObjectId oid)
        {
            oid = ObjectId.Null;
            if (db == null) return false;
            if (string.IsNullOrWhiteSpace(handleStr)) return false;

            try
            {
                long v = long.Parse(handleStr.Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var h = new Handle(v);
                oid = db.GetObjectId(false, h, 0);
                return oid != ObjectId.Null && oid.IsValid;
            }
            catch
            {
                return false;
            }
        }

        private static string CsvD(double? d)
            => d.HasValue ? d.Value.ToString("0.###", CultureInfo.InvariantCulture) : "";

        private static string CsvF(double d)
            => d.ToString("0.###", CultureInfo.InvariantCulture);


        private static string CsvLen4(double d)
            => Math.Round(d, 4).ToString("0.0000", CultureInfo.InvariantCulture);

        private static string CsvLen4(double? d)
            => d.HasValue ? Math.Round(d.Value, 4).ToString("0.0000", CultureInfo.InvariantCulture) : "";

        private static string CsvP(Point3d? p)
        {
            if (!p.HasValue) return ",,";
            return string.Join(",",
                p.Value.X.ToString("0.###", CultureInfo.InvariantCulture),
                p.Value.Y.ToString("0.###", CultureInfo.InvariantCulture),
                p.Value.Z.ToString("0.###", CultureInfo.InvariantCulture)
            );
        }

        // ----------------------------
        // State (QuantityIdUtil) helpers
        // ----------------------------
        private static int GetStateMaxId(QuantityIdUtil.State st, int fallback)
        {
            if (st == null) return fallback;
            try
            {
                var t = st.GetType();
                var p = t.GetProperty("MaxId", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                if (p != null)
                {
                    var v = p.GetValue(st, null);
                    if (v is int i) return i;
                    if (v != null && int.TryParse(v.ToString(), out var j)) return j;
                }
                var f = t.GetField("MaxId", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                if (f != null)
                {
                    var v = f.GetValue(st);
                    if (v is int i) return i;
                    if (v != null && int.TryParse(v.ToString(), out var j)) return j;
                }
            }
            catch { }
            return fallback;
        }


        // Natural sort comparator for work names like "STW 500 架設" < "STW 1000 架設"
        private static int CompareWorkNameNatural(string a, string b)
        {
            if (ReferenceEquals(a, b)) return 0;
            if (a == null) return -1;
            if (b == null) return 1;

            int ia = 0, ib = 0;
            while (ia < a.Length && ib < b.Length)
            {
                char ca = a[ia];
                char cb = b[ib];

                bool da = char.IsDigit(ca);
                bool db = char.IsDigit(cb);

                if (da && db)
                {
                    // parse consecutive digits as integers (ignore leading zeros)
                    long va = 0, vb = 0;
                    int sa = ia, sb = ib;

                    while (sa < a.Length && a[sa] == '0') sa++;
                    while (sb < b.Length && b[sb] == '0') sb++;

                    int ea = sa;
                    while (ea < a.Length && char.IsDigit(a[ea])) ea++;
                    int eb = sb;
                    while (eb < b.Length && char.IsDigit(b[eb])) eb++;

                    // length compare first to avoid overflow and keep numeric ordering
                    int lena = ea - sa;
                    int lenb = eb - sb;
                    if (lena != lenb) return lena.CompareTo(lenb);

                    // same digit length: compare value lexicographically
                    for (int k = 0; k < lena; k++)
                    {
                        int dcmp = a[sa + k].CompareTo(b[sb + k]);
                        if (dcmp != 0) return dcmp;
                    }

                    // numeric equal; tie-break by count of leading zeros (fewer zeros first)
                    int zca = sa - ia;
                    int zcb = sb - ib;
                    if (zca != zcb) return zca.CompareTo(zcb);

                    // advance original indices past full digit runs (including leading zeros)
                    while (ia < a.Length && char.IsDigit(a[ia])) ia++;
                    while (ib < b.Length && char.IsDigit(b[ib])) ib++;
                    continue;
                }

                // If one is digit and the other isn't, non-digit comes first (keeps "STW " grouping stable)
                if (da != db) return da ? 1 : -1;

                int c = ca.CompareTo(cb);
                if (c != 0) return c;

                ia++;
                ib++;
            }

            // shorter string first if all equal up to min length
            return a.Length.CompareTo(b.Length);
        }

        private static void SetStateMaxId(QuantityIdUtil.State st, int maxId)
        {
            if (st == null) return;
            try
            {
                var t = st.GetType();
                var p = t.GetProperty("MaxId", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                if (p != null && p.CanWrite)
                {
                    p.SetValue(st, maxId, null);
                    return;
                }
                var f = t.GetField("MaxId", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                if (f != null)
                {
                    f.SetValue(st, maxId);
                    return;
                }
            }
            catch { }
        }

        private static void SetStateNextId(QuantityIdUtil.State st, int nextId)
        {
            if (st == null) return;
            try
            {
                var t = st.GetType();
                // Some implementations use NextId, Next, or NextNumber
                foreach (var name in new[] { "NextId", "Next", "NextNumber" })
                {
                    var p = t.GetProperty(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                    if (p != null && p.CanWrite)
                    {
                        p.SetValue(st, nextId, null);
                        return;
                    }
                    var f = t.GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                    if (f != null)
                    {
                        f.SetValue(st, nextId);
                        return;
                    }
                }
            }
            catch { }
        }
    }
}
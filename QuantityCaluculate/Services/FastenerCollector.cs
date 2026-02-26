using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dObjects;
using Autodesk.AutoCAD.Geometry;

namespace UFlowPlant3D.Services
{
    /// <summary>
    /// FastenerRowId 収集ユーティリティ
    ///
    /// 重要:
    /// - Plant 3D のバージョン/参照DLL差で公開APIが揺れるため、
    ///   PnPDatabase/PnPRows/PnPDirectTableRow 等を“直接参照しない”実装（Reflection中心）。
    /// - PnPDatabase 側（Fasteners/Gasket/BoltSet/Buttweld）で取得した rowId は、
    ///   DWG(ModelSpace)に実体が無い“残骸”が混入することがあるため、
    ///   「DWGに存在する rowId」との積集合でフィルタして返す。
    /// </summary>
    public static class FastenerCollector
    {
        /// <summary>
        /// 推奨エントリポイント（MyCommands2 などから使用）
        /// </summary>
        public static List<int> CollectFastenerRowIds(Database db, DataLinksManager dlm, Editor ed = null)
        {
            // 現行方針：Connector（P3dConnector）を明細/集計から除外し、代わりに Gasket / BoltSet を「モデル内（現在DWG）」から収集して返す。
            //  - PnPDatabase の Fasteners テーブル全走査は “同一プロジェクト内の別DWG” まで混入することがあるため、ここでは原則使わない。
            //  - Gasket/BoltSet/Buttweld は Connector を起点に PnPDataLinks を引くことで、モデル内の実体RowId（Gasket/BoltSet/ButtweldテーブルのRowId）を得られる。
            var ids = new List<int>();
            if (db == null || dlm == null) return ids;

            var inst = CollectGasketBoltSetInstances(db, dlm, ed);
            if (inst != null && inst.Count > 0)
            {
                foreach (var x in inst)
                {
                    if (x == null) continue;
                    if (x.RowId > 0) ids.Add(x.RowId);
                }

                ids = ids.Distinct().ToList();
                try { ed?.WriteMessage($"\n[UFLOW] FastenerCollector: Gasket/BoltSet/Buttweld rowIds (in-dwg) = {ids.Count}"); } catch { }
                return ids;
            }

            // 何も取れなかった場合のみ、旧ロジックへフォールバック（※必要なら有効化して使う）
            // return CollectFastenerRowIds(dlm, ed);
            try { ed?.WriteMessage("\n[UFLOW][WARN] FastenerCollector: Gasket/BoltSet/Buttweld rowIds not found. (No fallback to Fasteners table scan)"); } catch { }
            return ids;
        }

        /// <summary>
        /// PnPDatabase の Fasteners/Gasket/BoltSet/Buttweld から rowId を収集（重複除去済み）
        /// ※この段階では“DWGに実体があるか”は見ていない（上位でフィルタする）
        /// </summary>
        public static List<int> CollectFastenerRowIds(DataLinksManager dlm, Editor ed = null)
        {
            // 互換用（旧呼び出し）。現在は “モデル内” 判定ができないため空を返す。
            // 必要なら呼び出し側で CollectFastenerRowIds(Database db, ...) を使うこと。
            try { ed?.WriteMessage("\n[UFLOW][WARN] FastenerCollector: CollectFastenerRowIds(dlm) is obsolete. Use CollectFastenerRowIds(db, dlm) instead."); } catch { }
            return new List<int>();
        }

        // ----------------------------
        // DWG(ModelSpace) 上に存在する rowId を全部集める（fastener判定なし）
        // ----------------------------
        private static HashSet<int> CollectRowIdsExistingInDrawing(Database db, DataLinksManager dlm, Editor ed)
        {
            var set = new HashSet<int>();
            if (db == null || dlm == null) return set;

            int scanned = 0, gotRowId = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId oid in ms)
                {
                    scanned++;
                    if (oid.IsNull || oid.IsErased) continue;

                    if (!TryFindRowIdFromObjectId(dlm, oid, out int rowId) || rowId <= 0)
                        continue;

                    gotRowId++;
                    set.Add(rowId);
                }

                tr.Commit();
            }

            ed?.WriteMessage($"\n[UFLOW][DBG] DWG rowId scan: scanned={scanned}, gotRowId={gotRowId}, uniqueRowId={set.Count}");
            return set;
        }

        /// <summary>
        /// 既存フォールバック：DWGのリンク付き図形からrowIdを収集
        /// </summary>
        private static List<int> CollectFastenerRowIdsByLinkedEntities(Database db, DataLinksManager dlm, Editor ed)
        {
            var set = new HashSet<int>();
            if (db == null || dlm == null) return set.OrderBy(x => x).ToList();

            int total = 0, hasRowId = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId oid in ms)
                {
                    total++;
                    if (oid.IsNull || oid.IsErased) continue;

                    if (!TryFindRowIdFromObjectId(dlm, oid, out int rowId) || rowId <= 0)
                        continue;

                    hasRowId++;
                    set.Add(rowId);
                }

                tr.Commit();
            }

            ed?.WriteMessage($"\n[UFLOW][DBG] ModelSpace total={total}, hasRowId={hasRowId}, uniqueRowIds={set.Count}");
            return set.OrderBy(x => x).ToList();
        }

        // ----------------------------
        // RowId取得（ObjectId -> RowId）
        // ----------------------------
        private static bool TryFindRowIdFromObjectId(DataLinksManager dlm, ObjectId oid, out int rowId)
        {
            rowId = 0;
            if (dlm == null || oid.IsNull) return false;

            string[] methodNames = new[]
            {
                "FindAcPpRowId",
                "FindRowId",
                "GetRowId"
            };

            foreach (var name in methodNames)
            {
                try
                {
                    var mi = dlm.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null,
                        new[] { typeof(ObjectId) }, null);
                    if (mi == null) continue;

                    var v = mi.Invoke(dlm, new object[] { oid });
                    if (v is int i && i > 0)
                    {
                        rowId = i;
                        return true;
                    }
                }
                catch { /* ignore */ }
            }
            return false;
        }

        // ----------------------------
        // PnPDatabase / Table / Rows 列挙（Reflection）
        // ----------------------------
        private static object TryGetPnPDatabase(DataLinksManager dlm)
        {
            if (dlm == null) return null;

            // 1) プロパティ
            try
            {
                var pi = dlm.GetType().GetProperty("PnPDatabase", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (pi != null)
                {
                    var v = pi.GetValue(dlm);
                    if (v != null) return v;
                }
            }
            catch { }

            // 2) メソッド
            foreach (var name in new[] { "GetPnPDatabase", "get_PnPDatabase" })
            {
                try
                {
                    var mi = dlm.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    if (mi == null) continue;
                    var v = mi.Invoke(dlm, null);
                    if (v != null) return v;
                }
                catch { }
            }

            return null;
        }

        private static object TryGetPnPTable(object pnpDb, string tableName, Editor ed = null)
        {
            if (pnpDb == null || string.IsNullOrWhiteSpace(tableName)) return null;

            // 0) Try direct indexer: pnpDb.Tables["Name"]
            try
            {
                var piTables = pnpDb.GetType().GetProperty("Tables", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                var tablesObj = piTables?.GetValue(pnpDb);
                if (tablesObj != null)
                {
                    var piItem = tablesObj.GetType().GetProperty("Item", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    if (piItem != null)
                    {
                        try
                        {
                            var t = piItem.GetValue(tablesObj, new object[] { tableName });
                            if (t != null) return t;
                        }
                        catch { /* some builds use int indexer */ }
                    }
                }
            }
            catch { }

            // 1) Try method: pnpDb.GetTable("Name")
            foreach (var m in new[] { "GetTable", "get_Table" })
            {
                try
                {
                    var mi = pnpDb.GetType().GetMethod(m, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null,
                        new[] { typeof(string) }, null);
                    if (mi == null) continue;
                    var t = mi.Invoke(pnpDb, new object[] { tableName });
                    if (t != null) return t;
                }
                catch { }
            }

            // 2) Enumerate tables and match by name (case-insensitive)
            try
            {
                var piTables = pnpDb.GetType().GetProperty("Tables", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                var tablesObj = piTables?.GetValue(pnpDb);
                if (tablesObj is IEnumerable en)
                {
                    foreach (var t in en)
                    {
                        if (t == null) continue;
                        string name = "";
                        try
                        {
                            var piName = t.GetType().GetProperty("Name", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                                         ?? t.GetType().GetProperty("TableName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                            name = piName != null ? (piName.GetValue(t)?.ToString() ?? "") : "";
                        }
                        catch { name = ""; }

                        if (name.Length == 0) continue;

                        if (string.Equals(name, tableName, StringComparison.OrdinalIgnoreCase))
                            return t;
                    }
                }
            }
            catch { }

            // 3) Last resort: dump available table names once (debug)
            try
            {
                var piTables = pnpDb.GetType().GetProperty("Tables", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                var tablesObj = piTables?.GetValue(pnpDb);
                if (ed != null && tablesObj is IEnumerable en)
                {
                    int n = 0;
                    ed.WriteMessage($"\n[UFLOW][DBG] PnPDatabase tables sample for lookup '{tableName}':");
                    foreach (var t in en)
                    {
                        if (t == null) continue;
                        string name = "";
                        try
                        {
                            var piName = t.GetType().GetProperty("Name", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                                         ?? t.GetType().GetProperty("TableName", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                            name = piName != null ? (piName.GetValue(t)?.ToString() ?? "") : "";
                        }
                        catch { name = ""; }
                        if (name.Length > 0)
                        {
                            ed.WriteMessage($"\n[UFLOW][DBG]   table[{++n}] {name}");
                            if (n >= 40) break;
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        private static IEnumerable EnumeratePnPTableRows(object table)
        {
            if (table == null) yield break;

            // 1) table.Select("1=1") を優先（戻り値が IEnumerable なら列挙）
            IEnumerable selected = TrySelectAllRows(table);
            if (selected != null)
            {
                foreach (var r in selected) yield return r;
                yield break;
            }

            // 2) table.Rows が IEnumerable なら列挙
            object rowsObj = null;
            try
            {
                var pi = table.GetType().GetProperty("Rows", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                rowsObj = pi?.GetValue(table);
            }
            catch { rowsObj = null; }

            if (rowsObj is IEnumerable en)
            {
                foreach (var r in en) yield return r;
                yield break;
            }

            // 3) どうしても取れない場合は何も返さない
            yield break;
        }

        private static IEnumerable TrySelectAllRows(object table)
        {
            try
            {
                var mi = table.GetType().GetMethod("Select", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null,
                    new[] { typeof(string) }, null);
                if (mi == null) return null;

                var ret = mi.Invoke(table, new object[] { "1=1" });
                return ret as IEnumerable;
            }
            catch
            {
                return null;
            }
        }

        private static int TryGetRowId(object row)
        {
            if (row == null) return 0;

            // RowId プロパティがあればそれ
            try
            {
                var pi = row.GetType().GetProperty("RowId", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (pi != null)
                {
                    var v = pi.GetValue(row);
                    if (v is int i) return i;
                    if (v != null && int.TryParse(v.ToString(), out var ii)) return ii;
                }
            }
            catch { }

            // PnPID など別名がある環境向け（念のため）
            foreach (var alt in new[] { "PnPID", "Id", "ID" })
            {
                try
                {
                    var pi = row.GetType().GetProperty(alt, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    if (pi == null) continue;
                    var v = pi.GetValue(row);
                    if (v is int i) return i;
                    if (v != null && int.TryParse(v.ToString(), out var ii)) return ii;
                }
                catch { }
            }

            return 0;
        }


        // ----------------------------
        // Gasket / BoltSet（Connector -> PnPDataLinks -> {Gasket,BoltSet} RowId）収集
        // ----------------------------

        public sealed class FastenerInstance
        {
            public int RowId { get; set; }
            public string ClassName { get; set; } = "";
            public string SourceConnectorHandle { get; set; } = "";
            public int DwgId { get; set; }
            public int DwgSubIndex { get; set; }
            public long DwgHandleLow { get; set; }
            public long DwgHandleHigh { get; set; }

            // Gasket の場合のみ（ConnectorのS1-S2を“厚み”として利用）
            public Point3d? S1 { get; set; }
            public Point3d? S2 { get; set; }
            public double ThicknessModelUnits { get; set; } = 0.0;
        }

        private sealed class PnPDataLinkRow
        {
            public int RowId;
            public int DwgId;
            public long HandleLow;
            public long HandleHigh;
            public int DwgSubIndex;
            public string RowClassName;
        }

        /// <summary>
        /// DWG上の Connector から PnPDataLinks を辿り、Gasket/BoltSet/Buttweld の rowId を回収する。
        /// - DwgHandleLow/High で束ね、baseDwgId（Connector本体のDwgId）に一致する行だけ採用。
        /// - 複数DWGの情報が混ざるケースがあるため、DwgIdフィルタが重要。
        /// </summary>
        public static List<FastenerInstance> CollectGasketBoltSetInstances(Database db, DataLinksManager dlm, Editor ed = null)
        {
            var list = new List<FastenerInstance>();
            if (db == null || dlm == null) return list;

            object pnpDb = TryGetPnPDatabase(dlm);
            if (pnpDb == null)
            {
                ed?.WriteMessage("\n[UFLOW] FastenerCollector: PnPDatabase not available (for PnPDataLinks query).");
                return list;
            }


            object pnpDataLinksTable = TryGetPnPTable(pnpDb, "PnPDataLinks", ed);
            if (pnpDataLinksTable == null)
            {
                ed?.WriteMessage("\n[UFLOW] FastenerCollector: PnPDataLinks table not found in PnPDatabase.");
                return list;
            }

            // NOTE:
            // Plant3D 2026 では PnPDataLinks の「全走査」がうまくいかない環境があるため、
            // Connector の baseRowId から Select(where) を使って必要行だけ引く（v13ログの方式）。
            var seen = new HashSet<int>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId oid in ms)
                {
                    if (oid.IsNull || oid.IsErased) continue;

                    var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    if (ent is not Connector) continue;

                    if (!TryFindRowIdFromObjectId(dlm, oid, out int baseRowId) || baseRowId <= 0)
                        continue;

                    // 1) baseRowId で PnPDataLinks を引く（where候補を順に試す）
                    IEnumerable baseRows = null;
                    string usedWhere = null;

                    foreach (var w in new[]
                    {
                $"RowId={baseRowId}",
                $"PnPID={baseRowId}",
                $"PnPId={baseRowId}",
                $"PnPRowId={baseRowId}"
            })
                    {
                        var rr = TrySelectRows(pnpDataLinksTable, w);
                        if (rr == null) continue;

                        var en = rr.GetEnumerator();
                        if (en != null && en.MoveNext())
                        {
                            baseRows = rr;
                            usedWhere = w;
                            break;
                        }
                    }

                    if (baseRows == null) continue;

                    // baseRow（Connector本体）を推定：RowClassName=P3dConnector & DwgSubIndex=0 を優先
                    int baseDwgId = 0;
                    long hLow = 0, hHigh = 0;

                    object fallbackFirst = null;
                    foreach (var r in baseRows)
                    {
                        if (r == null) continue;
                        if (fallbackFirst == null) fallbackFirst = r;

                        string cls = TryGetStringField(r, "RowClassName") ?? "";
                        int sub = TryGetIntField(r, "DwgSubIndex");
                        int dwgId = TryGetIntField(r, "DwgId");
                        long low = TryGetLongField(r, "DwgHandleLow");
                        long high = TryGetLongField(r, "DwgHandleHigh");

                        if (dwgId > 0 && baseDwgId == 0) baseDwgId = dwgId;

                        if (cls.IndexOf("P3dConnector", StringComparison.OrdinalIgnoreCase) >= 0 && sub == 0)
                        {
                            baseDwgId = dwgId;
                            hLow = low;
                            hHigh = high;
                            break;
                        }
                    }

                    // fallback
                    if (hLow == 0 && hHigh == 0 && fallbackFirst != null)
                    {
                        baseDwgId = TryGetIntField(fallbackFirst, "DwgId");
                        hLow = TryGetLongField(fallbackFirst, "DwgHandleLow");
                        hHigh = TryGetLongField(fallbackFirst, "DwgHandleHigh");
                    }

                    if (baseDwgId <= 0) continue;

                    // 2) handle で bundle を引く（AND が効かない環境もあるので候補を順に試す）
                    IEnumerable bundle = null;
                    foreach (var w in new[]
                    {
                $"DwgHandleLow={hLow} AND DwgHandleHigh={hHigh}",
                $"DwgHandleLow={hLow}"
            })
                    {
                        var rr = TrySelectRows(pnpDataLinksTable, w);
                        if (rr == null) continue;

                        var en = rr.GetEnumerator();
                        if (en != null && en.MoveNext())
                        {
                            bundle = rr;
                            break;
                        }
                    }

                    if (bundle == null) continue;

                    foreach (var br in bundle)
                    {
                        if (br == null) continue;

                        int dwgId = TryGetIntField(br, "DwgId");
                        if (dwgId != baseDwgId) continue; // 別DWG混入対策

                        string rc = (TryGetStringField(br, "RowClassName") ?? "").Trim();
                        if (!rc.Equals("Gasket", StringComparison.OrdinalIgnoreCase) &&
                            !rc.Equals("BoltSet", StringComparison.OrdinalIgnoreCase) &&
                            !rc.Equals("Buttweld", StringComparison.OrdinalIgnoreCase) &&
                            !rc.Equals("ButtWeld", StringComparison.OrdinalIgnoreCase))
                            continue;

                        int rid = TryGetIntColumn(br, "RowId"); // NOTE: PnPDataLinks has a RowId column (FK to real row). PnPRow.RowId may be the PnPID (PK), so use column-first access.
                        int linkId = TryGetIntField(br, "RowId"); // for debug only (may be PnPID)

                        if (rid <= 0 || !seen.Add(rid)) continue;

                        var inst = new FastenerInstance
                        {
                            RowId = rid,
                            ClassName = rc,
                            SourceConnectorHandle = ent.Handle.ToString(),
                            DwgId = dwgId,
                            DwgSubIndex = TryGetIntField(br, "DwgSubIndex"),
                            DwgHandleLow = TryGetLongField(br, "DwgHandleLow"),
                            DwgHandleHigh = TryGetLongField(br, "DwgHandleHigh")
                        };

                        // If linkId differs from rid, it usually means PnPRow.RowId (PK) != column RowId (FK to actual class table). We must use rid.
                        if (linkId > 0 && rid > 0 && linkId != rid)
                        {
                            // (suppressed) PnPDataLinks id mismatch log
                        }


                        // Gasketのみ：Connectorのports距離を“厚み”として利用
                        if (rc.Equals("Gasket", StringComparison.OrdinalIgnoreCase))
                        {
                            ComponentInfo ci = GeometryService.ExtractComponentInfo(dlm, oid, ent);
                            if (ci.Start.HasValue && ci.End.HasValue)
                            {
                                inst.S1 = ci.Start.Value;
                                inst.S2 = ci.End.Value;
                                inst.ThicknessModelUnits = ci.Start.Value.DistanceTo(ci.End.Value);
                            }
                        }

                        list.Add(inst);
                    }
                }

                tr.Commit();
            }

            ed?.WriteMessage($"\n[UFLOW] FastenerCollector: Gasket/BoltSet/Buttweld instances found={list.Count}");
            return list;
        }


        // ----------------------------
        // Connector bundle collection (for LineIndex)
        // ----------------------------

        public sealed class ConnectorBundle
        {
            public int DwgId { get; set; }
            public long DwgHandleLow { get; set; }
            public long DwgHandleHigh { get; set; }
            public string SourceConnectorHandle { get; set; } = "";
            public HashSet<int> RowIds { get; set; } = new HashSet<int>();
        }

        /// <summary>
        /// DWG上の Connector を起点に、PnPDataLinks の同一bundle（DwgHandleLow/High + DwgId）から
        /// filterRowIds に含まれる rowId（FK）を集めて返す。
        /// LineIndex（接続連結成分）の構築に使用する。
        /// </summary>
        public static List<ConnectorBundle> CollectConnectorBundles(Database db, DataLinksManager dlm, HashSet<int> filterRowIds, Editor ed = null)
        {
            var list = new List<ConnectorBundle>();
            if (db == null || dlm == null) return list;

            object pnpDb = TryGetPnPDatabase(dlm);
            if (pnpDb == null) return list;

            object pnpDataLinksTable = TryGetPnPTable(pnpDb, "PnPDataLinks", ed);
            if (pnpDataLinksTable == null) return list;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId oid in ms)
                {
                    if (oid.IsNull || oid.IsErased) continue;
                    var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                    if (ent is not Connector) continue;

                    if (!TryFindRowIdFromObjectId(dlm, oid, out int baseRowId) || baseRowId <= 0)
                        continue;

                    // 1) baseRowId で PnPDataLinks を引く（where候補を順に試す）
                    IEnumerable baseRows = null;
                    foreach (var w in new[]
                    {
                        $"RowId={baseRowId}",
                        $"PartRowId={baseRowId}",
                        $"ObjectId={baseRowId}"
                    })
                    {
                        var rr = TrySelectRows(pnpDataLinksTable, w);
                        if (rr == null) continue;
                        var en = rr.GetEnumerator();
                        if (en != null && en.MoveNext()) { baseRows = rr; break; }
                    }
                    if (baseRows == null) continue;

                    // baseRow（Connector本体）を推定：RowClassName=P3dConnector & DwgSubIndex=0 を優先
                    int baseDwgId = 0;
                    long hLow = 0, hHigh = 0;
                    object fallbackFirst = null;
                    foreach (var r in baseRows)
                    {
                        if (r == null) continue;
                        if (fallbackFirst == null) fallbackFirst = r;

                        string cls = TryGetStringField(r, "RowClassName") ?? "";
                        int sub = TryGetIntField(r, "DwgSubIndex");
                        int dwgId = TryGetIntField(r, "DwgId");
                        long low = TryGetLongField(r, "DwgHandleLow");
                        long high = TryGetLongField(r, "DwgHandleHigh");

                        if (dwgId > 0 && baseDwgId == 0) baseDwgId = dwgId;
                        if (cls.IndexOf("P3dConnector", StringComparison.OrdinalIgnoreCase) >= 0 && sub == 0)
                        {
                            baseDwgId = dwgId;
                            hLow = low;
                            hHigh = high;
                            break;
                        }
                    }

                    if (hLow == 0 && hHigh == 0 && fallbackFirst != null)
                    {
                        baseDwgId = TryGetIntField(fallbackFirst, "DwgId");
                        hLow = TryGetLongField(fallbackFirst, "DwgHandleLow");
                        hHigh = TryGetLongField(fallbackFirst, "DwgHandleHigh");
                    }

                    if (baseDwgId <= 0) continue;

                    // 2) handle で bundle を引く（AND が効かない環境もあるので候補を順に試す）
                    IEnumerable bundle = null;
                    foreach (var w in new[]
                    {
                        $"DwgHandleLow={hLow} AND DwgHandleHigh={hHigh}",
                        $"DwgHandleLow={hLow}"
                    })
                    {
                        var rr = TrySelectRows(pnpDataLinksTable, w);
                        if (rr == null) continue;
                        var en = rr.GetEnumerator();
                        if (en != null && en.MoveNext()) { bundle = rr; break; }
                    }
                    if (bundle == null) continue;

                    var ids = new HashSet<int>();
                    foreach (var br in bundle)
                    {
                        if (br == null) continue;
                        int dwgId = TryGetIntField(br, "DwgId");
                        if (dwgId != baseDwgId) continue; // 別DWG混入対策

                        int rid = TryGetIntColumn(br, "RowId");
                        if (rid <= 0) continue;
                        if (filterRowIds != null && filterRowIds.Count > 0 && !filterRowIds.Contains(rid)) continue;
                        ids.Add(rid);
                    }

                    if (ids.Count == 0) continue;

                    list.Add(new ConnectorBundle
                    {
                        DwgId = baseDwgId,
                        DwgHandleLow = hLow,
                        DwgHandleHigh = hHigh,
                        SourceConnectorHandle = ent.Handle.ToString(),
                        RowIds = ids
                    });
                }

                tr.Commit();
            }

            return list;
        }

        private static string MakeHandleKey(long low, long high) => $"{high}:{low}";

        private static int TryGetIntField(object row, string fieldName)
        {
            object v = TryGetField(row, fieldName);
            if (v == null) return 0;
            try
            {
                if (v is int i) return i;
                if (v is short s) return s;
                if (v is long l) return (l > int.MaxValue) ? 0 : (int)l;
                if (v is string str && int.TryParse(str, out int n)) return n;
            }
            catch { }
            return 0;
        }

        private static long TryGetLongField(object row, string fieldName)
        {
            object v = TryGetField(row, fieldName);
            if (v == null) return 0;
            try
            {
                if (v is long l) return l;
                if (v is int i) return i;
                if (v is short s) return s;
                if (v is string str && long.TryParse(str, out long n)) return n;
            }
            catch { }
            return 0;
        }

        private static string TryGetStringField(object row, string fieldName)
        {
            object v = TryGetField(row, fieldName);
            if (v == null) return null;
            try
            {
                return v.ToString();
            }
            catch { return null; }
        }

        private static object TryGetField(object row, string fieldName)
        {
            if (row == null || string.IsNullOrWhiteSpace(fieldName)) return null;
            var t = row.GetType();

            // 1) direct property
            try
            {
                var pi = t.GetProperty(fieldName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (pi != null) return pi.GetValue(row);
            }
            catch { }

            // 2) indexer: row["RowId"]
            try
            {
                var piItem = t.GetProperty("Item", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, null, new[] { typeof(string) }, null);
                if (piItem != null) return piItem.GetValue(row, new object[] { fieldName });
            }
            catch { }

            // 3) GetValue("RowId") 形式
            try
            {
                var mi = t.GetMethod("GetValue", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, new[] { typeof(string) }, null);
                if (mi != null) return mi.Invoke(row, new object[] { fieldName });
            }
            catch { }

            return null;
        }


        // Column-first access helper: for PnPRow, direct property RowId may be the row primary key (PnPID).
        // When we need a table column (e.g., PnPDataLinks.RowId = FK to actual class table), we must prefer indexer/GetValue.
        static object TryGetColumn(object row, string fieldName)
        {
            if (row == null || string.IsNullOrWhiteSpace(fieldName)) return null;
            var t = row.GetType();

            // 1) indexer: row["RowId"]
            try
            {
                var piItem = t.GetProperty("Item", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, null, new[] { typeof(string) }, null);
                if (piItem != null) return piItem.GetValue(row, new object[] { fieldName });
            }
            catch { }

            // 2) GetValue("RowId")
            try
            {
                var mi = t.GetMethod("GetValue", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, new[] { typeof(string) }, null);
                if (mi != null) return mi.Invoke(row, new object[] { fieldName });
            }
            catch { }

            // 3) direct property as last resort
            try
            {
                var pi = t.GetProperty(fieldName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (pi != null) return pi.GetValue(row);
            }
            catch { }

            return null;
        }

        static int TryGetIntColumn(object row, string fieldName)
        {
            object v = TryGetColumn(row, fieldName);
            if (v == null) return 0;
            try
            {
                if (v is int i) return i;
                if (v is short s) return s;
                if (v is long l) return (l > int.MaxValue) ? 0 : (int)l;
                if (v is string str && int.TryParse(str, out int n)) return n;
            }
            catch { }
            return 0;
        }

        private static IEnumerable TrySelectRows(object pnpTable, string where)
        {
            if (pnpTable == null) return null;

            try
            {
                var t = pnpTable.GetType();

                // Common signatures: Select(string), Select(string, string), Select(string, bool)
                var m1 = t.GetMethod("Select", new Type[] { typeof(string) });
                if (m1 != null) return m1.Invoke(pnpTable, new object[] { where }) as IEnumerable;

                var m2 = t.GetMethod("Select", new Type[] { typeof(string), typeof(string) });
                if (m2 != null) return m2.Invoke(pnpTable, new object[] { where, "" }) as IEnumerable;

                var m3 = t.GetMethod("Select", new Type[] { typeof(string), typeof(bool) });
                if (m3 != null) return m3.Invoke(pnpTable, new object[] { where, true }) as IEnumerable;
            }
            catch
            {
                // ignore
            }

            return null;
        }

    }
}
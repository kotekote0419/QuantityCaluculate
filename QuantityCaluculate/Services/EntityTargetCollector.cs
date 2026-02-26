using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFlowPlant3D.Services
{
    /// <summary>
    /// ModelSpace から「材料集計対象（Pipe/Part/PipeInlineAsset）」を収集する。
    /// 重要:
    /// - P3dConnector は ComponentList/Summary から除外する（材料ではなく接続表現のため）。
    /// - Gasket/BoltSet は FastenerCollector 側で Connector を走査して抽出する。
    /// </summary>
    public static class EntityTargetCollector
    {
        public static List<ObjectId> Collect(Database db)
        {
            var targets = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId oid in ms)
                {
                    if (oid.IsNull || oid.IsErased) continue;

                    var ent = tr.GetObject(oid, OpenMode.ForRead) as Entity;
                    if (ent == null) continue;

                    // Connector は除外（Gasket/BoltSet 抽出は別経路）
                    if (ent is Pipe || ent is Part || ent is PipeInlineAsset)
                        targets.Add(oid);
                }

                tr.Commit();
            }

            // 重複排除
            return new List<ObjectId>(new HashSet<ObjectId>(targets));
        }
    }
}

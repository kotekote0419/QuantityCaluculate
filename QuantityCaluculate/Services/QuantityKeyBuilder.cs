using System;
using System.Globalization;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFlowPlant3D.Services
{
    public static class QuantityKeyBuilder
    {
        public static string BuildKey(DataLinksManager dlm, ObjectId oid, Entity ent)
        {
            string install = PlantProp.GetString(dlm, oid, "施工方法", "Installation", "INSTALLATION");
            string size = PlantProp.GetString(dlm, oid, "Size", "サイズ", "NPS");

            if (ent is Pipe)
            {
                string mat = PlantProp.GetString(dlm, oid, "MaterialCode", "材料コード", "MAT_CODE");
                return JoinKey("PIPE", mat, install, size);
            }

            string pnpClass = PlantProp.GetString(dlm, oid, "PnPClassName", "ClassName");
            string kind = !string.IsNullOrWhiteSpace(pnpClass) ? pnpClass : ent.GetType().Name;

            string itemCode = PlantProp.GetString(dlm, oid, "ItemCode", "項目コード", "ITEM_CODE");
            string desc = PlantProp.GetString(dlm, oid, "PartFamilyLongDesc", "部品仕様詳細", "LONG_DESC", "ShortDescription");

            // 最低限：ASSET系でまとめる（必要ならここから種別キーを強化）
            return JoinKey(kind, Prefer(itemCode, desc), install, size);
        }

        // ★FastenerRowの直接集計キー
        public static string BuildFastenerKey(DataLinksManager dlm, int rowId)
        {
            string install = PlantProp.GetString(dlm, rowId, "施工方法", "Installation", "INSTALLATION");
            string size = PlantProp.GetString(dlm, rowId, "Size", "サイズ", "NPS");
            string itemCode = PlantProp.GetString(dlm, rowId, "ItemCode", "項目コード", "ITEM_CODE");
            string desc = PlantProp.GetString(dlm, rowId, "PartFamilyLongDesc", "部品仕様詳細", "LONG_DESC", "ShortDescription", "Description", "Desc");

            return JoinKey("FASTENER", Prefer(itemCode, desc), install, size);
        }

        private static string Prefer(string a, string b) => !string.IsNullOrWhiteSpace(a) ? a : (b ?? "");

        private static string JoinKey(params string[] parts)
        {
            for (int i = 0; i < parts.Length; i++)
                parts[i] = (parts[i] ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            return string.Join("|", parts);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.AutoCAD.DatabaseServices;

namespace UFlowPlant3D.Services
{
    /// <summary>
    /// DWGのNamedObjectsDictionaryに Key->ID を永続化（依存DLLなし）。
    /// 形式: 1行 = "<id>\t<key>"
    /// </summary>
    public sealed class QuantityIdStore
    {
        private const string DictKey = "UFLOW_QTYID_MAP_V1";

        public Dictionary<string, int> Map { get; set; } = new();
        public int MaxId { get; private set; } = 0;

        public static QuantityIdStore Load(Database db)
        {
            var store = new QuantityIdStore();

            using var tr = db.TransactionManager.StartTransaction();
            var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

            if (!nod.Contains(DictKey))
            {
                tr.Commit();
                return store;
            }

            var xrec = (Xrecord)tr.GetObject(nod.GetAt(DictKey), OpenMode.ForRead);

            string text = "";
            if (xrec.Data != null)
            {
                var arr = xrec.Data.AsArray();
                if (arr != null && arr.Length > 0 && arr[0].Value is string s)
                    text = s ?? "";
            }

            // parse
            // 1行: "<id>\t<key>"
            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var idx = line.IndexOf('\t');
                if (idx <= 0) continue;

                var idText = line.Substring(0, idx).Trim();
                var key = line.Substring(idx + 1);

                if (int.TryParse(idText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) && id > 0)
                {
                    if (!string.IsNullOrWhiteSpace(key))
                    {
                        store.Map[key] = id;
                        if (id > store.MaxId) store.MaxId = id;
                    }
                }
            }

            tr.Commit();
            return store;
        }

        public void Save(Database db)
        {
            using var tr = db.TransactionManager.StartTransaction();
            var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);

            // serialize: 1行 = "<id>\t<key>"
            // ※ key中に改行が入ると壊れるので、キー生成側では改行を入れない前提（通常OK）
            var lines = new List<string>(Map.Count);
            foreach (var kv in Map)
            {
                lines.Add($"{kv.Value.ToString(CultureInfo.InvariantCulture)}\t{kv.Key}");
            }
            var text = string.Join("\n", lines);

            Xrecord xrec;
            if (nod.Contains(DictKey))
            {
                xrec = (Xrecord)tr.GetObject(nod.GetAt(DictKey), OpenMode.ForWrite);
            }
            else
            {
                xrec = new Xrecord();
                nod.SetAt(DictKey, xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }

            xrec.Data = new ResultBuffer(new TypedValue((int)DxfCode.Text, text));
            tr.Commit();
        }

        public void ObserveExistingId(int id)
        {
            if (id > MaxId) MaxId = id;
        }

        public int GetOrCreate(string key)
        {
            if (Map.TryGetValue(key, out var id))
                return id;

            id = ++MaxId;
            Map[key] = id;
            return id;
        }
    }
}

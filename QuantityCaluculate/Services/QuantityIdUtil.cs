using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.AutoCAD.DatabaseServices;

namespace UFlowPlant3D.Services
{
    /// <summary>
    /// 集計キー文字列 -> 数量ID(数値) の対応を DWG に保存し、
    /// 「カウントアップ方式」で一意なIDを割り当てる。
    ///
    /// 必要なID数が N のとき、N の桁数を i とすると、
    /// 数量IDは i 桁（ゼロ埋め）で "000...1" ～ "N" を運用する。
    ///
    /// 保存先：NamedObjectsDictionary -> "UFLOW" Dictionary ->
    ///   Xrecord "QTYID_MAP"  : [Text key][Int32 id] のペア繰り返し
    ///   Xrecord "QTYID_NEXT" : 次に払い出すID(Int32)
    /// </summary>
    public static class QuantityIdUtil
    {
        private const string NOD_DICT = "UFLOW";
        private const string XREC_MAP = "QTYID_MAP";
        private const string XREC_NEXT = "QTYID_NEXT";

        public sealed class State
        {
            public Dictionary<string, int> Map { get; }
            public int Next { get; set; }
            public int Digits { get; }
            public int MaxId { get; }
            public int StartId { get; }

            public State(Dictionary<string, int> map, int next, int digits, int maxId, int startId)
            {
                Map = map;
                Next = next;
                Digits = digits;
                MaxId = maxId;
                StartId = startId;
            }
        }

        /// <summary>
        /// DWGからマップと次番号をロード。
        /// maxId=N のとき digits = 桁数(N) を自動採用（例: N=6000 -> 4桁）。
        /// </summary>
        public static State LoadState(Transaction tr, Database db, int maxId, int startId = 1)
        {
            if (maxId < 1) maxId = 1;
            if (startId < 1) startId = 1;

            int digits = GetDigitsFromMax(maxId);

            var map = LoadMap(tr, db);

            int next = LoadNext(tr, db)
                       ?? (map.Count == 0 ? startId : (map.Values.Max() + 1));

            if (next < startId) next = startId;

            return new State(map, next, digits, maxId, startId);
        }

        public static void SaveState(Transaction tr, Database db, State st)
        {
            if (st == null) return;
            SaveMap(tr, db, st.Map);
            SaveNext(tr, db, st.Next);
        }

        /// <summary>
        /// キーに対応するIDを返す。無ければ st.Next を割り当ててインクリメント。
        /// </summary>
        public static int GetOrCreateId(State st, string key)
        {
            if (st == null) throw new ArgumentNullException(nameof(st));
            key ??= "";

            if (st.Map.TryGetValue(key, out int existing)) return existing;

            if (st.Next > st.MaxId)
                throw new InvalidOperationException($"QuantityId exceeded MaxId={st.MaxId}. Next={st.Next}");

            int id = st.Next;
            st.Map[key] = id;
            st.Next++;
            return id;
        }

        public static string Format(State st, int id)
        {
            if (st == null) return id.ToString(CultureInfo.InvariantCulture);
            if (id < 0) return id.ToString(CultureInfo.InvariantCulture);
            return id.ToString(new string('0', st.Digits), CultureInfo.InvariantCulture);
        }

        private static int GetDigitsFromMax(int maxId)
        {
            // i = 桁数(N). 例: 9->1, 10->2, 6000->4
            return maxId.ToString(CultureInfo.InvariantCulture).Length;
        }

        // ----------------- NOD persistence -----------------
        private static Dictionary<string, int> LoadMap(Transaction tr, Database db)
        {
            var map = new Dictionary<string, int>(StringComparer.Ordinal);

            var xr = XrecInNod.GetOrNull(tr, db, NOD_DICT, XREC_MAP);
            if (xr?.Data == null) return map;

            var arr = xr.Data.AsArray();
            if (arr == null) return map;

            for (int i = 0; i + 1 < arr.Length; i += 2)
            {
                var k = arr[i].Value as string;
                if (string.IsNullOrEmpty(k)) continue;

                int id;
                try { id = Convert.ToInt32(arr[i + 1].Value, CultureInfo.InvariantCulture); }
                catch { continue; }

                if (!map.ContainsKey(k))
                    map.Add(k, id);
            }

            return map;
        }

        private static void SaveMap(Transaction tr, Database db, Dictionary<string, int> map)
        {
            if (map == null) return;

            var tvs = new List<TypedValue>(map.Count * 2);
            foreach (var kv in map.OrderBy(k => k.Value))
            {
                tvs.Add(new TypedValue((int)DxfCode.Text, kv.Key));
                tvs.Add(new TypedValue((int)DxfCode.Int32, kv.Value));
            }

            XrecInNod.Upsert(tr, db, NOD_DICT, XREC_MAP, new ResultBuffer(tvs.ToArray()));
        }

        private static int? LoadNext(Transaction tr, Database db)
        {
            var xr = XrecInNod.GetOrNull(tr, db, NOD_DICT, XREC_NEXT);
            if (xr?.Data == null) return null;

            var arr = xr.Data.AsArray();
            if (arr == null || arr.Length == 0) return null;

            try { return Convert.ToInt32(arr[0].Value, CultureInfo.InvariantCulture); }
            catch { return null; }
        }

        private static void SaveNext(Transaction tr, Database db, int next)
        {
            var rb = new ResultBuffer(new TypedValue((int)DxfCode.Int32, next));
            XrecInNod.Upsert(tr, db, NOD_DICT, XREC_NEXT, rb);
        }
    }

    internal static class XrecInNod
    {
        public static Xrecord GetOrNull(Transaction tr, Database db, string dictName, string xrecName)
        {
            if (tr == null || db == null) return null;

            var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
            if (nod == null || !nod.Contains(dictName)) return null;

            var dict = tr.GetObject(nod.GetAt(dictName), OpenMode.ForRead) as DBDictionary;
            if (dict == null || !dict.Contains(xrecName)) return null;

            return tr.GetObject(dict.GetAt(xrecName), OpenMode.ForRead) as Xrecord;
        }

        public static void Upsert(Transaction tr, Database db, string dictName, string xrecName, ResultBuffer rb)
        {
            var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);

            DBDictionary dict;
            if (nod.Contains(dictName))
            {
                dict = tr.GetObject(nod.GetAt(dictName), OpenMode.ForWrite) as DBDictionary;
            }
            else
            {
                dict = new DBDictionary();
                nod.SetAt(dictName, dict);
                tr.AddNewlyCreatedDBObject(dict, true);
            }

            Xrecord xr;
            if (dict.Contains(xrecName))
            {
                xr = tr.GetObject(dict.GetAt(xrecName), OpenMode.ForWrite) as Xrecord;
            }
            else
            {
                xr = new Xrecord();
                dict.SetAt(xrecName, xr);
                tr.AddNewlyCreatedDBObject(xr, true);
            }

            xr.Data = rb;
        }
    }
}

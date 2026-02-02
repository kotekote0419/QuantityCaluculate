using System;
using System.Collections;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;

namespace UFlowPlant3D.Services
{
    public static class PlantProp
    {
        public static string GetString(DataLinksManager dlm, ObjectId oid, params string[] candidates)
        {
            foreach (var name in candidates)
            {
                if (TryGet(dlm, oid, name, out var v))
                {
                    var s = v?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(s)) return s;
                }
            }
            return "";
        }

        public static double? GetDouble(DataLinksManager dlm, ObjectId oid, params string[] candidates)
        {
            foreach (var name in candidates)
            {
                if (TryGet(dlm, oid, name, out var v))
                {
                    if (v == null) continue;
                    if (v is double d) return d;
                    if (double.TryParse(v.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var dd)) return dd;
                }
            }
            return null;
        }

        // ---- public helper for other classes
        public static int? TryGetRowIdPublic(DataLinksManager dlm, ObjectId oid) => TryGetRowId(dlm, oid);

        public static object TryGetPpObjectId(DataLinksManager dlm, ObjectId oid)
        {
            // PpObjectId の取り方は環境差が大きいので反射で “あれば”
            try
            {
                var mi = dlm.GetType().GetMethod("GetPpObjectId", BindingFlags.Public | BindingFlags.Instance, null,
                    new[] { typeof(ObjectId) }, null);
                if (mi != null) return mi.Invoke(dlm, new object[] { oid });
            }
            catch { }
            return null;
        }

        // ---- core
        private static bool TryGet(DataLinksManager dlm, ObjectId oid, string propName, out object value)
        {
            value = null;

            int? rowId = TryGetRowId(dlm, oid);

            // RowId版
            if (rowId.HasValue)
            {
                if (TryGetFromProps(dlm, "GetAllProperties", new object[] { rowId.Value, true }, new[] { typeof(int), typeof(bool) }, propName, out value)) return true;
                if (TryGetFromProps(dlm, "GetProperties", new object[] { rowId.Value, true }, new[] { typeof(int), typeof(bool) }, propName, out value)) return true;
                if (TryGetValue(dlm, rowId.Value, propName, out value)) return value != null;
            }

            // ObjectId版
            if (TryGetFromProps(dlm, "GetAllProperties", new object[] { oid, true }, new[] { typeof(ObjectId), typeof(bool) }, propName, out value)) return true;
            if (TryGetFromProps(dlm, "GetProperties", new object[] { oid, true }, new[] { typeof(ObjectId), typeof(bool) }, propName, out value)) return true;
            if (TryGetValue(dlm, oid, propName, out value)) return value != null;

            return false;
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

        private static bool TryGetFromProps(DataLinksManager dlm, string methodName, object[] args, Type[] sig, string key, out object value)
        {
            value = null;
            try
            {
                var mi = dlm.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance, null, sig, null);
                if (mi == null) return false;

                var props = mi.Invoke(dlm, args);
                return TryReadFromProps(props, key, out value);
            }
            catch { return false; }
        }

        private static bool TryGetValue(DataLinksManager dlm, int rowId, string key, out object value)
        {
            value = null;
            try
            {
                var mi = dlm.GetType().GetMethod("GetPropertyValue", BindingFlags.Public | BindingFlags.Instance, null,
                    new[] { typeof(int), typeof(string) }, null);
                if (mi == null) return false;

                value = mi.Invoke(dlm, new object[] { rowId, key });
                return true;
            }
            catch { return false; }
        }

        private static bool TryGetValue(DataLinksManager dlm, ObjectId oid, string key, out object value)
        {
            value = null;
            try
            {
                var mi = dlm.GetType().GetMethod("GetPropertyValue", BindingFlags.Public | BindingFlags.Instance, null,
                    new[] { typeof(ObjectId), typeof(string) }, null);
                if (mi == null) return false;

                value = mi.Invoke(dlm, new object[] { oid, key });
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// GetAllProperties/GetProperties の戻り型差を吸収：
        /// - IDictionary / StringDictionary
        /// - NameValueCollection
        /// - IEnumerable（要素が Name/Value 等を持つ）
        /// </summary>
        private static bool TryReadFromProps(object props, string key, out object value)
        {
            value = null;
            if (props == null) return false;

            // IDictionary
            if (props is IDictionary dict)
            {
                if (dict.Contains(key))
                {
                    value = dict[key];
                    return true;
                }
                return false;
            }

            // NameValueCollection
            if (props is NameValueCollection nvc)
            {
                var s = nvc[key];
                if (s != null)
                {
                    value = s;
                    return true;
                }
                return false;
            }

            // IEnumerable: 各要素から名前と値を拾う
            if (props is IEnumerable e)
            {
                foreach (var item in e)
                {
                    if (item == null) continue;

                    var name = GetPropString(item, "Name")
                            ?? GetPropString(item, "Key")
                            ?? GetPropString(item, "PropertyName")
                            ?? GetPropString(item, "DisplayName");

                    if (!string.Equals(name, key, StringComparison.Ordinal)) continue;

                    value = GetPropObj(item, "Value")
                         ?? GetPropObj(item, "PropValue")
                         ?? GetPropObj(item, "PropertyValue");

                    // 値が取れない場合は item 自体
                    value ??= item.ToString();
                    return true;
                }
            }

            return false;
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
            return GetPropObj(obj, name)?.ToString();
        }
    }
}

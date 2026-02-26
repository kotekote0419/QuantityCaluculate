using System;
using System.Collections;
using System.Collections.Specialized;
using System.Globalization;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.DataLinks;

namespace UFlowPlant3D.Services
{
    public static class PlantProp
    {
        // ObjectId版
        public static string GetString(DataLinksManager dlm, ObjectId oid, params string[] candidates)
            => GetStringCore(dlm, oid, null, candidates);

        public static double? GetDouble(DataLinksManager dlm, ObjectId oid, params string[] candidates)
            => GetDoubleCore(dlm, oid, null, candidates);

        // rowId版（★FastenerRow用）
        public static string GetString(DataLinksManager dlm, int rowId, params string[] candidates)
            => GetStringCore(dlm, null, rowId, candidates);

        public static double? GetDouble(DataLinksManager dlm, int rowId, params string[] candidates)
            => GetDoubleCore(dlm, null, rowId, candidates);

        private static string GetStringCore(DataLinksManager dlm, ObjectId? oid, int? rowId, params string[] candidates)
        {
            foreach (var name in candidates)
            {
                if (TryGet(dlm, oid, rowId, name, out var v))
                {
                    var s = v?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(s)) return s;
                }
            }
            return "";
        }

        private static double? GetDoubleCore(DataLinksManager dlm, ObjectId? oid, int? rowId, params string[] candidates)
        {
            foreach (var name in candidates)
            {
                if (TryGet(dlm, oid, rowId, name, out var v))
                {
                    if (v == null) continue;
                    if (v is double d) return d;
                    if (double.TryParse(v.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var dd)) return dd;
                }
            }
            return null;
        }

        private static bool TryGet(DataLinksManager dlm, ObjectId? oid, int? rowId, string propName, out object value)
        {
            value = null;
            if (dlm == null || string.IsNullOrWhiteSpace(propName)) return false;

            // rowId優先
            if (rowId.HasValue && rowId.Value > 0)
            {
                if (TryGetFromProps(dlm, "GetAllProperties", new object[] { rowId.Value, true }, new[] { typeof(int), typeof(bool) }, propName, out value)) return true;
                if (TryGetFromProps(dlm, "GetProperties", new object[] { rowId.Value, true }, new[] { typeof(int), typeof(bool) }, propName, out value)) return true;
                if (TryGetValue(dlm, rowId.Value, propName, out value)) return value != null;
                return false;
            }

            if (!oid.HasValue) return false;

            if (TryGetFromProps(dlm, "GetAllProperties", new object[] { oid.Value, true }, new[] { typeof(ObjectId), typeof(bool) }, propName, out value)) return true;
            if (TryGetFromProps(dlm, "GetProperties", new object[] { oid.Value, true }, new[] { typeof(ObjectId), typeof(bool) }, propName, out value)) return true;
            if (TryGetValue(dlm, oid.Value, propName, out value)) return value != null;

            return false;
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

        // ★大小文字無視で読む
        private static bool TryReadFromProps(object props, string key, out object value)
        {
            value = null;
            if (props == null || string.IsNullOrWhiteSpace(key)) return false;

            if (props is IDictionary dict)
            {
                foreach (DictionaryEntry de in dict)
                {
                    var k = de.Key?.ToString();
                    if (k == null) continue;
                    if (string.Equals(k, key, StringComparison.OrdinalIgnoreCase))
                    {
                        value = de.Value;
                        return true;
                    }
                }
                return false;
            }

            if (props is NameValueCollection nvc)
            {
                foreach (var k in nvc.AllKeys)
                {
                    if (k == null) continue;
                    if (string.Equals(k, key, StringComparison.OrdinalIgnoreCase))
                    {
                        value = nvc[k];
                        return value != null;
                    }
                }
                return false;
            }

            if (props is IEnumerable e)
            {
                foreach (var item in e)
                {
                    if (item == null) continue;

                    var name = GetPropString(item, "Name")
                            ?? GetPropString(item, "Key")
                            ?? GetPropString(item, "PropertyName")
                            ?? GetPropString(item, "DisplayName");

                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (!string.Equals(name, key, StringComparison.OrdinalIgnoreCase)) continue;

                    value = GetPropObj(item, "Value")
                         ?? GetPropObj(item, "PropValue")
                         ?? GetPropObj(item, "PropertyValue");

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

        private static string GetPropString(object obj, string name) => GetPropObj(obj, name)?.ToString();
    }
}

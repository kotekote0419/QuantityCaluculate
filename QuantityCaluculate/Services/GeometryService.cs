using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.PnP3dObjects;

namespace UFlowPlant3D.Services
{
    public sealed class ComponentInfo
    {
        public string HandleString { get; set; } = "";
        public string EntityType { get; set; } = "";

        public double? ND1 { get; set; }
        public double? ND2 { get; set; }
        public double? ND3 { get; set; }

        public Point3d? Start { get; set; }
        public Point3d? End { get; set; }
        public Point3d? Branch { get; set; }
        public Point3d? Branch2 { get; set; }
        public Point3d? Mid { get; set; }
    }

    public static class GeometryService
    {
        /// <summary>
        /// Class4.cs の思想を踏襲して座標情報を抽出
        /// - Pipe: StartPoint/EndPoint を reflection で取得、ND1 は PartSizeProperties.PropValue("NominalDiameter")
        /// - それ以外: GetPorts(PortType.Static) があれば反射で呼び、Port index 0/1/2/3 を採用
        /// - Mid: Valve系はCOG、その他はPosition、無ければ中点
        ///
        /// ※ Autodesk.ProcessPower.PnPConnectors に依存しない
        /// </summary>
        public static ComponentInfo ExtractComponentInfo(DataLinksManager dlm, ObjectId oid, Entity ent)
        {
            var info = new ComponentInfo
            {
                HandleString = oid.Handle.ToString(),
                EntityType = ent.GetType().Name
            };

            // Pipe
            if (ent is Pipe pipe)
            {
                info.ND1 = TryGetPipeNominalDiameter(pipe);

                info.Start = TryGetPointByReflection(pipe, "StartPoint");
                info.End = TryGetPointByReflection(pipe, "EndPoint");

                info.Mid = ComputeMidFromPropsOrFallback(dlm, oid, info.Start, info.End, info.EntityType);
                return info;
            }

            // Pipe以外：GetPorts(PortType.Static) を持っていれば ports 抽出
            var portsObj = TryInvokeGetPorts(ent, PortType.Static);
            if (portsObj != null)
            {
                FillPortsFromStatic(portsObj, info);
            }

            info.Mid = ComputeMidFromPropsOrFallback(dlm, oid, info.Start, info.End, info.EntityType);
            return info;
        }

        // ---- Ports（戻り型差を吸収）
        private static void FillPortsFromStatic(object portsObj, ComponentInfo info)
        {
            var ports = PortUtil.ToPortArray(portsObj);

            // Port index固定：0=Start,1=End,2=Branch,3=Branch2（あなたの仕様）
            if (ports.Length >= 1) info.Start = ports[0].Position;
            if (ports.Length >= 2) info.End = ports[1].Position;
            if (ports.Length >= 3) info.Branch = ports[2].Position;
            if (ports.Length >= 4) info.Branch2 = ports[3].Position;

            // ND（取れる環境なら）
            try
            {
                if (ports.Length >= 1) info.ND1 = ToDoubleMaybe(ports[0].NominalDiameter);
                if (ports.Length >= 3) info.ND2 = ToDoubleMaybe(ports[2].NominalDiameter);
                if (ports.Length >= 4) info.ND3 = ToDoubleMaybe(ports[3].NominalDiameter);
            }
            catch { /* ignore */ }
        }

        private static double? ToDoubleMaybe(object v)
        {
            if (v == null) return null;
            if (v is double d) return d;
            if (double.TryParse(v.ToString(), out var dd)) return dd;
            return null;
        }

        // ---- Midpoint
        private static Point3d? ComputeMidFromPropsOrFallback(DataLinksManager dlm, ObjectId oid, Point3d? s, Point3d? e, string typeName)
        {
            // Valve系はCOG優先（Class4.csの思想）
            if (typeName.IndexOf("Valve", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var cog = TryGetPointFromProps(dlm, oid, "COG X", "COG Y", "COG Z");
                if (cog.HasValue) return cog;
            }

            // それ以外はPosition
            var pos = TryGetPointFromProps(dlm, oid, "Position X", "Position Y", "Position Z");
            if (pos.HasValue) return pos;

            // fallback：中点
            if (s.HasValue && e.HasValue)
                return s.Value + (e.Value - s.Value) * 0.5;

            return null;
        }

        private static Point3d? TryGetPointFromProps(DataLinksManager dlm, ObjectId oid, string xName, string yName, string zName)
        {
            var x = PlantProp.GetDouble(dlm, oid, xName);
            var y = PlantProp.GetDouble(dlm, oid, yName);
            var z = PlantProp.GetDouble(dlm, oid, zName);
            if (x.HasValue && y.HasValue && z.HasValue)
                return new Point3d(x.Value, y.Value, z.Value);
            return null;
        }

        // ---- Pipe ND
        private static double? TryGetPipeNominalDiameter(Pipe pipe)
        {
            try
            {
                var psp = pipe.PartSizeProperties;
                if (psp != null)
                {
                    var v = psp.PropValue("NominalDiameter");
                    if (v is double d) return d;
                    if (v != null && double.TryParse(v.ToString(), out var dd)) return dd;
                }
            }
            catch { }
            return null;
        }

        private static Point3d? TryGetPointByReflection(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) return null;
                var v = pi.GetValue(obj);
                if (v is Point3d p) return p;
            }
            catch { }
            return null;
        }

        /// <summary>
        /// ent.GetPorts(PortType) を反射で呼ぶ（Connector/Part 等の名前空間差を吸収）
        /// </summary>
        private static object TryInvokeGetPorts(object ent, PortType portType)
        {
            try
            {
                // GetPorts(PortType) を探索
                var mi = ent.GetType().GetMethod("GetPorts", BindingFlags.Public | BindingFlags.Instance, null, new[] { typeof(PortType) }, null);
                if (mi == null) return null;

                return mi.Invoke(ent, new object[] { portType });
            }
            catch
            {
                return null;
            }
        }

        private static class PortUtil
        {
            public static Port[] ToPortArray(object portsObj)
            {
                if (portsObj == null) return Array.Empty<Port>();

                // PortCollection等は IEnumerable を実装していることが多い
                if (portsObj is IEnumerable e)
                {
                    var list = new List<Port>();
                    foreach (var x in e)
                        if (x is Port p) list.Add(p);
                    return list.ToArray();
                }

                return Array.Empty<Port>();
            }
        }
    }
}

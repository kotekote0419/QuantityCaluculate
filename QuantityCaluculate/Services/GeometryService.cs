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
    public static class GeometryService
    {
        /// <summary>
        /// Class4.cs の思想を踏襲して座標情報を抽出
        /// - Pipe: StartPoint/EndPoint を reflection で取得、ND1 は PartSizeProperties.PropValue("NominalDiameter")
        /// - それ以外: GetPorts(PortType.Static) があれば反射で呼び、Port index 0/1/2/3 を採用
        /// - Mid: Elbow は Center、Tee は BranchPoint を優先し、無ければCOG/Position/中点
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

                info.Mid = ComputeMidFromPropsOrFallback(dlm, oid, info.Start, info.End, info.EntityType, ent);
                return info;
            }

            // Pipe以外：GetPorts(PortType.Static) を持っていれば ports 抽出
            var portsObj = TryInvokeGetPorts(ent, PortType.Static);
            if (portsObj != null)
            {
                FillPortsFromStatic(portsObj, info);
            }

            info.Mid = ComputeMidFromPropsOrFallback(dlm, oid, info.Start, info.End, info.EntityType, ent);
            return info;
        }

        // ---- Ports（戻り型差を吸収）
        private static void FillPortsFromStatic(object portsObj, ComponentInfo info)
        {
            var ports = PortUtil.ToPortObjects(portsObj);

            // Port index固定：0=Start,1=End,2=Branch,3=Branch2（仕様）
            if (ports.Count >= 1) info.Start = TryGetPointByReflection(ports[0], "Position");
            if (ports.Count >= 2) info.End = TryGetPointByReflection(ports[1], "Position");
            if (ports.Count >= 3) info.Branch = TryGetPointByReflection(ports[2], "Position");
            if (ports.Count >= 4) info.Branch2 = TryGetPointByReflection(ports[3], "Position");

            // ND（取れる環境なら）
            if (ports.Count >= 1) info.ND1 = ToDoubleMaybe(TryGetPropertyValue(ports[0], "NominalDiameter"));
            if (ports.Count >= 3) info.ND2 = ToDoubleMaybe(TryGetPropertyValue(ports[2], "NominalDiameter"));
            if (ports.Count >= 4) info.ND3 = ToDoubleMaybe(TryGetPropertyValue(ports[3], "NominalDiameter"));
        }

        private static double? ToDoubleMaybe(object v)
        {
            if (v == null) return null;
            if (v is double d) return d;
            if (double.TryParse(v.ToString(), out var dd)) return dd;
            return null;
        }

        // ---- Midpoint / Center
        private static Point3d? ComputeMidFromPropsOrFallback(DataLinksManager dlm, ObjectId oid, Point3d? s, Point3d? e, string typeName, object ent)
        {
            // Elbow: Center or CenterPoint
            if (typeName.IndexOf("Elbow", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var center = TryGetPointByReflection(ent, "Center")
                             ?? TryGetPointByReflection(ent, "CenterPoint")
                             ?? TryGetPointFromProps(dlm, oid, "Center X", "Center Y", "Center Z");
                if (center.HasValue) return center;
            }

            // Tee: BranchPoint
            if (typeName.IndexOf("Tee", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var branchPoint = TryGetPointByReflection(ent, "BranchPoint")
                                  ?? TryGetPointFromProps(dlm, oid, "BranchPoint X", "BranchPoint Y", "BranchPoint Z");
                if (branchPoint.HasValue) return branchPoint;
            }

            // Valve系はCOG優先
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

        private static object TryGetPropertyValue(object obj, string propName)
        {
            try
            {
                var pi = obj.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                return pi?.GetValue(obj);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// ent.GetPorts(PortType) を反射で呼ぶ（Connector/Part 等の名前空間差を吸収）
        /// </summary>
        private static object TryInvokeGetPorts(object ent, PortType portType)
        {
            try
            {
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
            public static List<object> ToPortObjects(object portsObj)
            {
                var list = new List<object>();
                if (portsObj == null) return list;

                if (portsObj is IEnumerable e)
                {
                    foreach (var x in e)
                    {
                        if (x != null) list.Add(x);
                    }
                }

                return list;
            }
        }
    }
}

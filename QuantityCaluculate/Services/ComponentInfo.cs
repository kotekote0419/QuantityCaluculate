using Autodesk.AutoCAD.Geometry;

namespace UFlowPlant3D.Services
{
    public class ComponentInfo
    {
        public string HandleString { get; set; } = "";
        public string EntityType { get; set; } = "";

        public double? ND1 { get; set; }
        public double? ND2 { get; set; }
        public double? ND3 { get; set; }

        public Point3d? Start { get; set; }
        public Point3d? Mid { get; set; }
        public Point3d? End { get; set; }

        public Point3d? Branch { get; set; }
        public Point3d? Branch2 { get; set; }
    }
}

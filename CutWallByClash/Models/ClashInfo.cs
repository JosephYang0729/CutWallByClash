using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace CutWallByClash.Models
{
    public class ClashInfo
    {
        public Wall Wall { get; set; }
        public Element MEPElement { get; set; }
        public XYZ ClashPoint { get; set; }
        public MEPCategory Category { get; set; }
        public BoundingBoxXYZ MEPBoundingBox { get; set; }
        public double ElementDiameter { get; set; }
        public double ElementWidth { get; set; }
        public double ElementHeight { get; set; }
    }

    public class OpeningGroup
    {
        public List<ClashInfo> Clashes { get; set; } = new List<ClashInfo>();
        public XYZ CenterPoint { get; set; }
        public Wall Wall { get; set; }
        public double RequiredWidth { get; set; }
        public double RequiredHeight { get; set; }
        public double RequiredDiameter { get; set; }
        public OpeningType OpeningType { get; set; }
    }
}
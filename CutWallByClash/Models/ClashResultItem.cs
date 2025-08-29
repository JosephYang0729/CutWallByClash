using Autodesk.Revit.DB;

namespace CutWallByClash.Models
{
    public class ClashResultItem
    {
        public int Index { get; set; }
        public string WallId { get; set; }
        public string WallName { get; set; }
        public string MEPType { get; set; }
        public string MEPId { get; set; }
        public string ElementSize { get; set; }
        public string OpeningId { get; set; } // 新增開口ID屬性
        
        // 保留原始對象引用用於3D定位
        public ClashInfo OriginalClash { get; set; }
    }
}
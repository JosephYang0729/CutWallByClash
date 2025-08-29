using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;

namespace CutWallByClash.Core
{
    public static class ViewChecker
    {
        /// <summary>
        /// 檢查當前視圖是否為3D視圖，如果不是則嘗試切換
        /// </summary>
        /// <param name="uiDoc">UI文檔</param>
        /// <returns>是否成功確保在3D視圖中</returns>
        public static bool EnsureIn3DView(UIDocument uiDoc)
        {
            var doc = uiDoc.Document;
            var activeView = doc.ActiveView;

            // 如果已經在3D視圖中，直接返回true
            if (activeView.ViewType == ViewType.ThreeD)
            {
                return true;
            }

            // 嘗試找到3D視圖
            var view3D = Get3DView(doc);
            if (view3D == null)
            {
                return false;
            }

            // 切換到3D視圖
            try
            {
                uiDoc.ActiveView = view3D;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 獲取文檔中的3D視圖
        /// </summary>
        /// <param name="document">文檔</param>
        /// <returns>3D視圖，如果找不到則返回null</returns>
        public static View3D Get3DView(Document document)
        {
            var collector = new FilteredElementCollector(document)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate);

            // 優先選擇名稱包含"3D"的視圖
            var view3D = collector.FirstOrDefault(v => v.Name.Contains("3D")) ?? collector.FirstOrDefault();

            return view3D;
        }

        /// <summary>
        /// 檢查當前視圖類型並返回描述
        /// </summary>
        /// <param name="view">視圖</param>
        /// <returns>視圖類型描述</returns>
        public static string GetViewTypeDescription(View view)
        {
            switch (view.ViewType)
            {
                case ViewType.ThreeD:
                    return "3D視圖";
                case ViewType.FloorPlan:
                    return "平面圖";
                case ViewType.CeilingPlan:
                    return "天花板平面圖";
                case ViewType.Elevation:
                    return "立面圖";
                case ViewType.Section:
                    return "剖面圖";
                case ViewType.Detail:
                    return "詳圖";
                case ViewType.Schedule:
                    return "明細表";
                case ViewType.DrawingSheet:
                    return "圖紙";
                case ViewType.Report:
                    return "報告";
                case ViewType.DraftingView:
                    return "製圖視圖";
                case ViewType.Legend:
                    return "圖例";
                case ViewType.EngineeringPlan:
                    return "工程平面圖";
                case ViewType.AreaPlan:
                    return "面積平面圖";
                case ViewType.Rendering:
                    return "彩現";
                default:
                    return "未知視圖類型";
            }
        }
    }
}
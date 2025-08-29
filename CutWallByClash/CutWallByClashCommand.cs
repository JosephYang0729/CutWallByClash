using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CutWallByClash.UI;
using System;
using System.Linq;

namespace CutWallByClash
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CutWallByClashCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc.Document;

                if (doc == null)
                {
                    message = "沒有開啟的Revit文檔。";
                    return Result.Failed;
                }

                // 檢查是否為3D視圖，如果不是則自動切換
                var activeView = doc.ActiveView;
                if (activeView.ViewType != ViewType.ThreeD)
                {
                    // 嘗試切換到3D視圖
                    var view3D = Get3DView(doc);
                    if (view3D != null)
                    {
                        uiDoc.ActiveView = view3D;
                        //TaskDialog.Show("資訊", "已自動切換到3D視圖以獲得最佳效果。");
                    }
                    else
                    {
                        //TaskDialog.Show("警告", "找不到3D視圖，建議手動切換到3D視圖後再執行此工具。");
                    }
                }

                // 開啟主視窗
                var mainWindow = new MainWindow(doc);
                mainWindow.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"執行命令時發生錯誤：{ex.Message}";
                return Result.Failed;
            }
        }

        private View3D Get3DView(Document document)
        {
            var collector = new FilteredElementCollector(document)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate);

            // 優先選擇名稱包含"3D"的視圖
            var view3D = collector.FirstOrDefault(v => v.Name.Contains("3D")) ?? collector.FirstOrDefault();

            return view3D;
        }
    }
}
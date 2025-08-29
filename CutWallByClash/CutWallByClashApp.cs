using Autodesk.Revit.UI;
using System;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace CutWallByClash
{
    public class CutWallByClashApp : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // 創建Ribbon面板
                var ribbonPanel = application.CreateRibbonPanel("牆體開口工具");

                // 創建按鈕
                var assemblyPath = Assembly.GetExecutingAssembly().Location;
                var buttonData = new PushButtonData(
                    "CutWallByClash",
                    "牆體開口\n建立工具",
                    assemblyPath,
                    "CutWallByClash.CutWallByClashCommand");

                buttonData.ToolTip = "根據MEP元件與牆的碰撞自動建立開口";
                buttonData.LongDescription = "此工具可以偵測專案中所有牆與MEP元件（管道、風管、電纜架、電管）的碰撞，並自動建立相應的開口。支援矩形和圓形開口，並可合併相近的碰撞點。";

                // 設定圖示（可選）
                try
                {
                    var iconUri = new Uri("pack://application:,,,/CutWallByClash;component/Resources/icon32.png");
                    buttonData.LargeImage = new BitmapImage(iconUri);
                }
                catch
                {
                    // 如果沒有圖示檔案，忽略錯誤
                }

                var pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("錯誤", $"載入外掛程式時發生錯誤：{ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
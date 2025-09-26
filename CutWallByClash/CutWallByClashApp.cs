using Autodesk.Revit.UI;
using System;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace CutWallByClash
{
    public class CutWallByClashApp : IExternalApplication
    {
        static string addinAssmeblyPath = Assembly.GetExecutingAssembly().Location;
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                try { application.CreateRibbonTab("中興航空城"); } catch { }
                RibbonPanel ribbonPanel = application.GetRibbonPanels("中興航空城").FirstOrDefault(p => p.Name == "小工具");
                if (ribbonPanel == null)
                {
                    ribbonPanel = application.CreateRibbonPanel("中興航空城", "小工具");
                }

                PushButton pushbutton1 = ribbonPanel.AddItem(
                new PushButtonData("CutWallByClash", "牆面自動開口",
                    addinAssmeblyPath, "CutWallByClash.CutWallByClashCommand"))
                        as PushButton;
                pushbutton1.ToolTip = "根據MEP元件與牆的碰撞位置與尺寸，自動建立相對應大小及形狀的開口元件";
                pushbutton1.LargeImage = new BitmapImage(new Uri("pack://application:,,,/CutWallByClash;component/Resources/CutWallByClash.png"));

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
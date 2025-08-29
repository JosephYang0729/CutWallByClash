using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CutWallByClash.Core;
using CutWallByClash.Models;

namespace CutWallByClash.UI
{
    public partial class MainWindow : Window
    {
        private readonly Document _document;
        private readonly WallOpeningEventHandler _eventHandler;
        private List<ClashInfo> _detectedClashes;
        private Dictionary<ClashInfo, ElementId> _openingIds; // 存儲碰撞和對應開口ID的映射
        
        // 保存矩形和圓形開口的參數值
        private double _savedRectWidth = 50;    // 矩形開口距離(H)預設值
        private double _savedRectHeight = 100;  // 矩形開口距離(V)預設值
        private double _savedCircularDiameter = 12.7; // 圓形開口距離預設值

        public MainWindow(Document document)
        {
            InitializeComponent();
            _document = document;
            _eventHandler = new WallOpeningEventHandler();
            _detectedClashes = new List<ClashInfo>();
            _openingIds = new Dictionary<ClashInfo, ElementId>();
            
            // 初始化介面狀態
            InitializeUI();
        }
        
        private void InitializeUI()
        {
            // 延遲初始化，確保所有UI元件都已載入
            this.Loaded += (s, e) =>
            {
                // 設定初始的開口類型顯示
                if (pnlRectangular != null && pnlCircular != null)
                {
                    pnlRectangular.Visibility = System.Windows.Visibility.Visible;
                    pnlCircular.Visibility = System.Windows.Visibility.Collapsed;
                }
                
                // 設定初始圖片顯示
                if (imgRectangular != null) imgRectangular.Visibility = System.Windows.Visibility.Visible;
                if (imgCircular != null) imgCircular.Visibility = System.Windows.Visibility.Collapsed;
            };
        }

        private void OnOpeningTypeChanged(object sender, RoutedEventArgs e)
        {
            if (pnlRectangular == null || pnlCircular == null) return;
            
            if (rbRectangular?.IsChecked == true)
            {
                // 保存圓形參數值
                SaveCircularParameters();
                
                // 顯示矩形參數
                pnlRectangular.Visibility = System.Windows.Visibility.Visible;
                pnlCircular.Visibility = System.Windows.Visibility.Collapsed;
                
                // 載入矩形參數值
                LoadRectangularParameters();
                
                // 切換圖片顯示
                if (imgRectangular != null) imgRectangular.Visibility = System.Windows.Visibility.Visible;
                if (imgCircular != null) imgCircular.Visibility = System.Windows.Visibility.Collapsed;
            }
            else if (rbCircular?.IsChecked == true)
            {
                // 保存矩形參數值
                SaveRectangularParameters();
                
                // 顯示圓形參數
                pnlRectangular.Visibility = System.Windows.Visibility.Collapsed;
                pnlCircular.Visibility = System.Windows.Visibility.Visible;
                
                // 載入圓形參數值
                LoadCircularParameters();
                
                // 切換圖片顯示
                if (imgRectangular != null) imgRectangular.Visibility = System.Windows.Visibility.Collapsed;
                if (imgCircular != null) imgCircular.Visibility = System.Windows.Visibility.Visible;
            }
        }

        private void SaveRectangularParameters()
        {
            if (txtRectWidth != null && double.TryParse(txtRectWidth.Text, out double width) && width > 0)
            {
                _savedRectWidth = width;
            }
            if (txtRectHeight != null && double.TryParse(txtRectHeight.Text, out double height) && height > 0)
            {
                _savedRectHeight = height;
            }
        }

        private void SaveCircularParameters()
        {
            if (txtCircularDiameter != null && double.TryParse(txtCircularDiameter.Text, out double diameter) && diameter > 0)
            {
                _savedCircularDiameter = diameter;
            }
        }

        private void LoadRectangularParameters()
        {
            if (txtRectWidth != null) txtRectWidth.Text = _savedRectWidth.ToString();
            if (txtRectHeight != null) txtRectHeight.Text = _savedRectHeight.ToString();
        }

        private void LoadCircularParameters()
        {
            if (txtCircularDiameter != null) txtCircularDiameter.Text = _savedCircularDiameter.ToString();
        }

        private async void OnExecute(object sender, RoutedEventArgs e)
        {
            var startTime = DateTime.Now;
            
            try
            {
                btnExecute.IsEnabled = false;
                txtStatus.Text = "正在偵測碰撞...";
                progressBar.Value = 0;

                var selectedCategories = GetSelectedCategories();
                if (!selectedCategories.Any())
                {
                    MessageBox.Show("請至少選擇一種MEP元件類型進行偵測。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var parameters = GetOpeningParameters();
                if (parameters == null) return;

                // 第一階段：碰撞偵測
                var progress = new Progress<int>(value =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        progressBar.Value = value * 0.5; // 偵測佔50%進度
                        txtStatus.Text = $"正在偵測碰撞... {value}%";
                    });
                });

                _detectedClashes = await Task.Run(() =>
                {
                    try
                    {
                        var clashDetector = new ClashDetector(_document);
                        return clashDetector.DetectClashes(selectedCategories, progress);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"碰撞偵測失敗: {ex.Message}", ex);
                    }
                });

                if (!_detectedClashes.Any())
                {
                    txtStatus.Text = "未發現碰撞點";
                    MessageBox.Show("未發現任何碰撞點。", "資訊", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 第二階段：建立開口
                txtStatus.Text = "正在建立開口...";
                
                var openingProgress = new Progress<int>(value =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        progressBar.Value = 50 + (value * 0.5); // 開口建立佔50%進度
                        txtStatus.Text = $"正在建立開口... {value}%";
                    });
                });

                _eventHandler.ExecuteOpeningCreation(
                    _detectedClashes,
                    parameters,
                    openingProgress,
                    (message, openingIds) => OnExecutionCompleted(message, startTime, openingIds),
                    (error) => OnExecutionError(error)
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"執行過程中發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                txtStatus.Text = "執行失敗";
                ResetUI();
            }
        }

        private void OnExecutionCompleted(string message, DateTime startTime, Dictionary<ClashInfo, ElementId> openingIds)
        {
            Dispatcher.Invoke(() =>
            {
                var executionTime = DateTime.Now - startTime;
                txtStatus.Text = "執行完成";
                progressBar.Value = 100;
                
                // 儲存開口ID映射
                _openingIds = openingIds;
                
                // 創建結果數據
                var results = CreateResultItems();
                var summary = $"成功建立 {results.Count} 個開口，偵測到 {_detectedClashes.Count} 個碰撞點";
                
                // 顯示結果窗口
                var resultsWindow = new ResultsWindow(results, summary, executionTime, _document);
                resultsWindow.Show();
                
                // 關閉當前窗口
                this.Close();
            });
        }

        private void OnExecutionError(string error)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = "執行失敗";
                progressBar.Value = 0;
                MessageBox.Show(error, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                ResetUI();
            });
        }

        private void ResetUI()
        {
            btnExecute.IsEnabled = true;
            progressBar.Value = 0;
        }

        private HashSet<MEPCategory> GetSelectedCategories()
        {
            var categories = new HashSet<MEPCategory>();
            
            if (chkPipes.IsChecked == true) categories.Add(MEPCategory.Pipes);
            if (chkDucts.IsChecked == true) categories.Add(MEPCategory.Ducts);
            if (chkCableTray.IsChecked == true) categories.Add(MEPCategory.CableTray);
            if (chkConduit.IsChecked == true) categories.Add(MEPCategory.Conduit);
            if (chkFlexPipe.IsChecked == true) categories.Add(MEPCategory.FlexPipe);
            
            return categories;
        }

        private OpeningParameters GetOpeningParameters()
        {
            try
            {
                // 先保存當前顯示的參數
                if (rbRectangular?.IsChecked == true)
                {
                    SaveRectangularParameters();
                }
                else if (rbCircular?.IsChecked == true)
                {
                    SaveCircularParameters();
                }

                var parameters = new OpeningParameters();
                
                // 智能開口系統不需要指定開口類型，會自動根據MEP元件類型決定
                // 但我們仍然需要設定所有參數值供系統使用
                
                // 設定矩形開口參數（使用保存的值）
                parameters.RectangularWidth = _savedRectWidth;
                parameters.RectangularHeight = _savedRectHeight;
                
                // 設定圓形開口參數（使用保存的值）
                parameters.CircularDiameter = _savedCircularDiameter;

                // 設定預設值
                parameters.WallThickness = 50; // 預設牆厚餘量
                parameters.MergeDistance = 100; // 預設合併距離

                return parameters;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"參數解析錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private ObservableCollection<ClashResultItem> CreateResultItems()
        {
            var results = new ObservableCollection<ClashResultItem>();
            
            for (int i = 0; i < _detectedClashes.Count; i++)
            {
                var clash = _detectedClashes[i];
                var openingId = _openingIds.ContainsKey(clash) ? _openingIds[clash].ToString() : "未建立";
                
                var resultItem = new ClashResultItem
                {
                    Index = i + 1,
                    WallId = clash.Wall.Id.ToString(),
                    WallName = clash.Wall.Name ?? "未命名牆",
                    MEPType = GetCategoryDisplayName(clash.Category),
                    MEPId = clash.MEPElement.Id.ToString(),
                    ElementSize = GetElementSizeString(clash),
                    OpeningId = openingId,
                    OriginalClash = clash
                };
                
                results.Add(resultItem);
            }
            
            return results;
        }

        private string GetElementSizeString(ClashInfo clash)
        {
            if (clash.ElementDiameter > 0)
            {
                return $"Ø{Math.Round(clash.ElementDiameter, 1)}mm";
            }
            else if (clash.ElementWidth > 0 && clash.ElementHeight > 0)
            {
                return $"{Math.Round(clash.ElementWidth, 1)}×{Math.Round(clash.ElementHeight, 1)}mm";
            }
            return "未知尺寸";
        }

        private string GetCategoryDisplayName(MEPCategory category)
        {
            switch (category)
            {
                case MEPCategory.Pipes:
                    return "管道";
                case MEPCategory.Ducts:
                    return "風管";
                case MEPCategory.CableTray:
                    return "電纜架";
                case MEPCategory.Conduit:
                    return "電管";
                case MEPCategory.FlexPipe:
                    return "撓性管";
                default:
                    return category.ToString();
            }
        }

    }
}
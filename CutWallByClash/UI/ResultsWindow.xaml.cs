using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CutWallByClash.Core;
using CutWallByClash.Models;

namespace CutWallByClash.UI
{
    public partial class ResultsWindow : Window
    {
        private readonly View3DNavigationHandler _navigationHandler;
        private ObservableCollection<ClashResultItem> _results;
        private readonly Document _document;
        private ClashResultItem _currentEditingItem;
        private readonly OpeningEditExternalEvent _editExternalEvent;
        private readonly ExternalEvent _editRevitEvent;
        private readonly CustomNavigationAndSelectionEvent _customNavigationEvent;
        private readonly ExternalEvent _customNavigationRevitEvent;

        public ResultsWindow(ObservableCollection<ClashResultItem> results, string summary, TimeSpan executionTime, Document document = null)
        {
            InitializeComponent();
            _navigationHandler = new View3DNavigationHandler();
            _results = results;
            _document = document;
            _currentEditingItem = null;
            
            // 初始化ExternalEvent
            _editExternalEvent = new OpeningEditExternalEvent();
            _editRevitEvent = ExternalEvent.Create(_editExternalEvent);
            
            // 初始化自定義導航ExternalEvent
            _customNavigationEvent = new CustomNavigationAndSelectionEvent();
            _customNavigationRevitEvent = ExternalEvent.Create(_customNavigationEvent);
            
            // 設定數據源
            dgResults.ItemsSource = _results;
            
            // 設定統計信息
            txtResultsSummary.Text = summary;
            txtExecutionTime.Text = $"執行時間: {executionTime.TotalSeconds:F1} 秒";
            
            // 初始化編輯面板
            InitializeEditPanel();
        }

        private void InitializeEditPanel()
        {
            // 初始設定為矩形顯示
            pnlRectangularEdit.Visibility = System.Windows.Visibility.Visible;
            pnlCircularEdit.Visibility = System.Windows.Visibility.Collapsed;
            imgEditRectangular.Visibility = System.Windows.Visibility.Visible;
            imgEditCircular.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void OnResultsDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // 檢查是否點擊在按鈕上，如果是則不處理
            var originalSource = e.OriginalSource as FrameworkElement;
            if (originalSource is Button || IsChildOfButton(originalSource))
            {
                e.Handled = true;
                return;
            }

            // 獲取點擊位置對應的行
            var hitTest = VisualTreeHelper.HitTest(dgResults, e.GetPosition(dgResults));
            if (hitTest?.VisualHit != null)
            {
                var row = FindParent<DataGridRow>(hitTest.VisualHit);
                if (row?.Item is ClashResultItem selectedItem && selectedItem.OriginalClash != null)
                {
                    try
                    {
                        // 設置選中項
                        dgResults.SelectedItem = selectedItem;
                        
                        // 直接執行我們的導航和選取邏輯，不使用原本的 NavigateToClash
                        NavigateAndSelectWallAndOpening(selectedItem);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"定位元件時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        private bool IsChildOfButton(FrameworkElement element)
        {
            if (element == null) return false;
            
            var parent = element.Parent as FrameworkElement;
            while (parent != null)
            {
                if (parent is Button) return true;
                parent = parent.Parent as FrameworkElement;
            }
            return false;
        }

        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            var parent = VisualTreeHelper.GetParent(child);
            while (parent != null)
            {
                if (parent is T result)
                    return result;
                parent = VisualTreeHelper.GetParent(parent);
            }
            return null;
        }

        private void NavigateAndSelectWallAndOpening(ClashResultItem selectedItem)
        {
            try
            {
                if (_document == null) return;
                
                var elementsToSelect = new List<ElementId>();
                
                // 只添加牆體（確保是當前文檔中的牆體，不是連結模型）
                if (selectedItem.OriginalClash?.Wall != null)
                {
                    var wall = selectedItem.OriginalClash.Wall;
                    // 確認牆體屬於當前文檔
                    if (wall.Document.Equals(_document))
                    {
                        elementsToSelect.Add(wall.Id);
                    }
                }

                // 只添加開口元件（如果有開口ID）
                if (!string.IsNullOrEmpty(selectedItem.OpeningId) && selectedItem.OpeningId != "未建立")
                {
                    if (TryParseElementId(selectedItem.OpeningId, out ElementId openingId))
                    {
                        var openingElement = _document.GetElement(openingId);
                        // 確認開口元件存在且屬於當前文檔
                        if (openingElement != null && openingElement.Document.Equals(_document))
                        {
                            elementsToSelect.Add(openingId);
                        }
                    }
                }

                // 設置導航和選取數據
                _customNavigationEvent.SetData(
                    selectedItem.OriginalClash,
                    elementsToSelect,
                    (message) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            // 可以在這裡添加成功提示，或者保持靜默
                        });
                    },
                    (error) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show(error, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                    }
                );

                _customNavigationRevitEvent.Raise();
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"導航和選取元件時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private void SelectWallAndOpening(ClashResultItem selectedItem)
        {
            try
            {
                if (_document == null) return;

                // 使用 ExternalEvent 來執行選取操作
                var selectionEvent = new SelectionExternalEvent();
                var revitEvent = ExternalEvent.Create(selectionEvent);
                
                var elementsToSelect = new List<ElementId>();
                
                // 只添加牆體（確保是當前文檔中的牆體，不是連結模型）
                if (selectedItem.OriginalClash?.Wall != null)
                {
                    var wall = selectedItem.OriginalClash.Wall;
                    // 確認牆體屬於當前文檔
                    if (wall.Document.Equals(_document))
                    {
                        elementsToSelect.Add(wall.Id);
                    }
                }

                // 只添加開口元件（如果有開口ID）
                if (!string.IsNullOrEmpty(selectedItem.OpeningId) && selectedItem.OpeningId != "未建立")
                {
                    if (TryParseElementId(selectedItem.OpeningId, out ElementId openingId))
                    {
                        var openingElement = _document.GetElement(openingId);
                        // 確認開口元件存在且屬於當前文檔
                        if (openingElement != null && openingElement.Document.Equals(_document))
                        {
                            elementsToSelect.Add(openingId);
                        }
                    }
                }

                // 執行選取（明確清除之前的選取，只選取指定元件）
                if (elementsToSelect.Any())
                {
                    selectionEvent.SetElementIds(elementsToSelect);
                    revitEvent.Raise();
                }
                else
                {
                    // 如果沒有要選取的元件，清空選取
                    selectionEvent.SetElementIds(new List<ElementId>());
                    revitEvent.Raise();
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"選取元件時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private bool TryParseElementId(string idString, out ElementId elementId)
        {
            elementId = ElementId.InvalidElementId;
            
            if (string.IsNullOrEmpty(idString))
                return false;
                
            try
            {
                if (int.TryParse(idString, out int id))
                {
                    elementId = new ElementId(id);
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private void OnModifyOpening(object sender, RoutedEventArgs e)
        {
            try
            {
                var button = sender as Button;
                var item = button?.Tag as ClashResultItem;
                
                if (item == null || string.IsNullOrEmpty(item.OpeningId) || item.OpeningId == "未建立")
                {
                    MessageBox.Show("此項目沒有對應的開口元件可以修改。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _currentEditingItem = item;
                LoadOpeningParameters(item);
                ShowEditPanel();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入開口參數時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnDeleteOpening(object sender, RoutedEventArgs e)
        {
            try
            {
                var button = sender as Button;
                var item = button?.Tag as ClashResultItem;
                
                if (item == null || string.IsNullOrEmpty(item.OpeningId) || item.OpeningId == "未建立")
                {
                    MessageBox.Show("此項目沒有對應的開口元件可以刪除。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var result = MessageBox.Show($"確定要刪除開口 ID: {item.OpeningId} 嗎？\n此操作無法復原。", 
                    "確認刪除", MessageBoxButton.YesNo, MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    DeleteOpening(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刪除開口時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadOpeningParameters(ClashResultItem item)
        {
            if (_document == null) return;

            if (TryParseElementId(item.OpeningId, out ElementId openingId))
            {
                var openingElement = _document.GetElement(openingId) as FamilyInstance;
                if (openingElement != null)
                {
                    var familySymbol = openingElement.Symbol;
                    bool isRectangular = familySymbol.Name.Contains("矩形");
                    
                    if (isRectangular)
                    {
                        // 矩形開口
                        lblOpeningType.Content = "矩形開口";
                        pnlRectangularEdit.Visibility = System.Windows.Visibility.Visible;
                        pnlCircularEdit.Visibility = System.Windows.Visibility.Collapsed;
                        imgEditRectangular.Visibility = System.Windows.Visibility.Visible;
                        imgEditCircular.Visibility = System.Windows.Visibility.Collapsed;

                        // 載入參數值
                        var widthParam = openingElement.LookupParameter("開口寬度");
                        var heightParam = openingElement.LookupParameter("開口長度");
                        var elevationParam = openingElement.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);

                        if (widthParam != null)
                            txtEditRectWidth.Text = Math.Round(widthParam.AsDouble() * 304.8, 1).ToString(); // 轉換為mm
                        if (heightParam != null)
                            txtEditRectHeight.Text = Math.Round(heightParam.AsDouble() * 304.8, 1).ToString();
                        if (elevationParam != null)
                            txtEditRectElevation.Text = Math.Round(elevationParam.AsDouble() * 304.8 / 1000, 3).ToString(); // 轉換為m
                    }
                    else
                    {
                        // 圓形開口
                        lblOpeningType.Content = "圓形開口";
                        pnlRectangularEdit.Visibility = System.Windows.Visibility.Collapsed;
                        pnlCircularEdit.Visibility = System.Windows.Visibility.Visible;
                        imgEditRectangular.Visibility = System.Windows.Visibility.Collapsed;
                        imgEditCircular.Visibility = System.Windows.Visibility.Visible;

                        // 載入參數值
                        var diameterParam = openingElement.LookupParameter("直徑");
                        var elevationParam = openingElement.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);

                        if (diameterParam != null)
                            txtEditCircularDiameter.Text = Math.Round(diameterParam.AsDouble() * 304.8, 1).ToString();
                        if (elevationParam != null)
                            txtEditCircularElevation.Text = Math.Round(elevationParam.AsDouble() * 304.8 / 1000, 3).ToString(); // 轉換為m
                    }
                }
            }
        }

        private void ShowEditPanel()
        {
            grpOpeningEdit.Visibility = System.Windows.Visibility.Visible;
            
            // 滾動到編輯面板
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var scrollViewer = FindVisualChild<ScrollViewer>(this);
                scrollViewer?.ScrollToBottom();
            }), System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void DeleteOpening(ClashResultItem item)
        {
            if (_document == null) return;

            if (TryParseElementId(item.OpeningId, out ElementId openingId))
            {
                _editExternalEvent.SetDeleteOperation(
                    openingId,
                    item,
                    (message) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            // 從表格中移除該項目
                            _results.Remove(item);
                            MessageBox.Show(message, "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                        });
                    },
                    (error) =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show(error, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                    }
                );
                _editRevitEvent.Raise();
            }
        }

        private void OnSaveChanges(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentEditingItem == null) return;

                SaveOpeningParameters(_currentEditingItem);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"儲存參數時發生錯誤：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnCancelEdit(object sender, RoutedEventArgs e)
        {
            HideEditPanel();
        }

        private void SaveOpeningParameters(ClashResultItem item)
        {
            if (_document == null) return;

            if (TryParseElementId(item.OpeningId, out ElementId openingId))
            {
                var openingElement = _document.GetElement(openingId) as FamilyInstance;
                if (openingElement != null)
                {
                    var familySymbol = openingElement.Symbol;
                    bool isRectangular = familySymbol.Name.Contains("矩形");
                    var parameters = new Dictionary<string, double>();

                    if (isRectangular)
                    {
                        // 矩形開口參數
                        if (double.TryParse(txtEditRectWidth.Text, out double width))
                            parameters["width"] = width;

                        if (double.TryParse(txtEditRectHeight.Text, out double height))
                            parameters["height"] = height;

                        if (double.TryParse(txtEditRectElevation.Text, out double elevation))
                            parameters["elevation"] = elevation * 1000; // 從m轉換為mm
                    }
                    else
                    {
                        // 圓形開口參數
                        if (double.TryParse(txtEditCircularDiameter.Text, out double diameter))
                            parameters["diameter"] = diameter;

                        if (double.TryParse(txtEditCircularElevation.Text, out double elevation))
                            parameters["elevation"] = elevation * 1000; // 從m轉換為mm
                    }

                    _editExternalEvent.SetModifyOperation(
                        openingId,
                        parameters,
                        (message) =>
                        {
                            Dispatcher.Invoke(() =>
                            {
                                HideEditPanel();
                                MessageBox.Show(message, "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                            });
                        },
                        (error) =>
                        {
                            Dispatcher.Invoke(() =>
                            {
                                MessageBox.Show(error, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                            });
                        }
                    );
                    _editRevitEvent.Raise();
                }
            }
        }

        private void HideEditPanel()
        {
            grpOpeningEdit.Visibility = System.Windows.Visibility.Collapsed;
            _currentEditingItem = null;
        }

        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T result)
                    return result;
                
                var childOfChild = FindVisualChild<T>(child);
                if (childOfChild != null)
                    return childOfChild;
            }
            return null;
        }

        private void OnClose(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }

    public enum EditOperation
    {
        Delete,
        Modify
    }

    public class OpeningEditExternalEvent : IExternalEventHandler
    {
        private EditOperation _operation;
        private ElementId _openingId;
        private ClashResultItem _item;
        private Dictionary<string, double> _parameters;
        private Action<string> _onCompleted;
        private Action<string> _onError;

        public void SetDeleteOperation(ElementId openingId, ClashResultItem item, Action<string> onCompleted, Action<string> onError)
        {
            _operation = EditOperation.Delete;
            _openingId = openingId;
            _item = item;
            _onCompleted = onCompleted;
            _onError = onError;
        }

        public void SetModifyOperation(ElementId openingId, Dictionary<string, double> parameters, Action<string> onCompleted, Action<string> onError)
        {
            _operation = EditOperation.Modify;
            _openingId = openingId;
            _parameters = parameters;
            _onCompleted = onCompleted;
            _onError = onError;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var document = app.ActiveUIDocument.Document;

                using (var transaction = new Transaction(document, _operation == EditOperation.Delete ? "Delete Opening" : "Modify Opening"))
                {
                    transaction.Start();
                    try
                    {
                        if (_operation == EditOperation.Delete)
                        {
                            document.Delete(_openingId);
                            _onCompleted?.Invoke("開口已成功刪除。");
                        }
                        else if (_operation == EditOperation.Modify)
                        {
                            var openingElement = document.GetElement(_openingId) as FamilyInstance;
                            if (openingElement != null)
                            {
                                var familySymbol = openingElement.Symbol;
                                bool isRectangular = familySymbol.Name.Contains("矩形");

                                if (isRectangular)
                                {
                                    if (_parameters.ContainsKey("width"))
                                    {
                                        var widthParam = openingElement.LookupParameter("開口寬度");
                                        widthParam?.Set(_parameters["width"] / 304.8);
                                    }
                                    if (_parameters.ContainsKey("height"))
                                    {
                                        var heightParam = openingElement.LookupParameter("開口長度");
                                        heightParam?.Set(_parameters["height"] / 304.8);
                                    }
                                    if (_parameters.ContainsKey("elevation"))
                                    {
                                        var elevationParam = openingElement.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                                        elevationParam?.Set(_parameters["elevation"] / 304.8);
                                    }
                                }
                                else
                                {
                                    if (_parameters.ContainsKey("diameter"))
                                    {
                                        var diameterParam = openingElement.LookupParameter("直徑");
                                        diameterParam?.Set(_parameters["diameter"] / 304.8);
                                    }
                                    if (_parameters.ContainsKey("elevation"))
                                    {
                                        var elevationParam = openingElement.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM);
                                        elevationParam?.Set(_parameters["elevation"] / 304.8);
                                    }
                                }
                            }
                            _onCompleted?.Invoke("開口參數已成功更新。");
                        }

                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.RollBack();
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                _onError?.Invoke($"操作失敗：{ex.Message}");
            }
        }

        public string GetName()
        {
            return "Opening Edit External Event";
        }
    }

    public class SelectionExternalEvent : IExternalEventHandler
    {
        private List<ElementId> _elementIds;

        public void SetElementIds(List<ElementId> elementIds)
        {
            _elementIds = elementIds;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                if (_elementIds != null && _elementIds.Any())
                {
                    var uiDoc = app.ActiveUIDocument;
                    uiDoc.Selection.SetElementIds(_elementIds);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("錯誤", $"選取元件時發生錯誤：{ex.Message}");
            }
        }

        public string GetName()
        {
            return "Selection External Event";
        }
    }

    public class CustomNavigationAndSelectionEvent : IExternalEventHandler
    {
        private ClashInfo _clashInfo;
        private List<ElementId> _elementsToSelect;
        private Action<string> _onCompleted;
        private Action<string> _onError;

        public void SetData(ClashInfo clashInfo, List<ElementId> elementsToSelect, Action<string> onCompleted, Action<string> onError)
        {
            _clashInfo = clashInfo;
            _elementsToSelect = elementsToSelect;
            _onCompleted = onCompleted;
            _onError = onError;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var uiDoc = app.ActiveUIDocument;
                var document = uiDoc.Document;

                // 確保在3D視圖中
                var view3D = Get3DView(document);
                if (view3D == null)
                {
                    _onError?.Invoke("找不到3D視圖，請確保專案中有3D視圖。");
                    return;
                }

                // 切換到3D視圖
                uiDoc.ActiveView = view3D;

                // 設置3D剖面框
                using (var transaction = new Transaction(document, "Navigate to Clash"))
                {
                    transaction.Start();

                    try
                    {
                        // 計算剖面框範圍（以碰撞點為中心，擴展5米）
                        var clashPoint = _clashInfo.ClashPoint;
                        var boxSize = 5.0 / 0.3048; // 5米轉換為英尺

                        var min = new XYZ(
                            clashPoint.X - boxSize,
                            clashPoint.Y - boxSize,
                            clashPoint.Z - boxSize
                        );

                        var max = new XYZ(
                            clashPoint.X + boxSize,
                            clashPoint.Y + boxSize,
                            clashPoint.Z + boxSize
                        );

                        var boundingBox = new BoundingBoxXYZ
                        {
                            Min = min,
                            Max = max
                        };

                        // 設置剖面框
                        view3D.SetSectionBox(boundingBox);

                        transaction.Commit();

                        // 刷新視圖
                        uiDoc.RefreshActiveView();

                        // 只選取指定的元件（牆壁和開口）
                        if (_elementsToSelect != null && _elementsToSelect.Any())
                        {
                            uiDoc.Selection.SetElementIds(_elementsToSelect);
                        }
                        else
                        {
                            // 清空選取
                            uiDoc.Selection.SetElementIds(new List<ElementId>());
                        }

                        // 縮放到適合
                        uiDoc.GetOpenUIViews().FirstOrDefault()?.ZoomToFit();

                        _onCompleted?.Invoke("成功定位到碰撞點並選取牆壁和開口元件");
                    }
                    catch (Exception ex)
                    {
                        transaction.RollBack();
                        _onError?.Invoke($"設置3D剖面框失敗: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _onError?.Invoke($"導航到碰撞點時發生錯誤：{ex.Message}");
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

        public string GetName()
        {
            return "Custom Navigation and Selection Event";
        }
    }
}
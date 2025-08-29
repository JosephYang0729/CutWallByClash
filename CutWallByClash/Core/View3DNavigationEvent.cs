using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using CutWallByClash.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CutWallByClash.Core
{
    public class View3DNavigationEvent : IExternalEventHandler
    {
        private ClashInfo _clashInfo;
        private Action<string> _onCompleted;
        private Action<string> _onError;

        public void SetData(ClashInfo clashInfo, Action<string> onCompleted, Action<string> onError)
        {
            _clashInfo = clashInfo;
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
                using (var transaction = new Transaction(document, "Set Section Box"))
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

                        // 選取碰撞的牆壁和MEP元件，以及連結的管線
                        var elementIds = new List<ElementId>();
                        var linkInstanceIds = new List<ElementId>();
                        
                        // 處理牆的選取
                        if (IsElementInCurrentDocument(_clashInfo.Wall, document))
                        {
                            elementIds.Add(_clashInfo.Wall.Id);
                        }
                        else
                        {
                            // 連結模型的牆，需要選取連結實例
                            var linkInstance = GetLinkInstanceForElement(_clashInfo.Wall, document);
                            if (linkInstance != null)
                            {
                                linkInstanceIds.Add(linkInstance.Id);
                            }
                        }
                        
                        // 處理MEP元件的選取
                        if (IsElementInCurrentDocument(_clashInfo.MEPElement, document))
                        {
                            elementIds.Add(_clashInfo.MEPElement.Id);
                        }
                        else
                        {
                            // 連結模型的MEP元件，需要選取連結實例
                            var linkInstance = GetLinkInstanceForElement(_clashInfo.MEPElement, document);
                            if (linkInstance != null)
                            {
                                linkInstanceIds.Add(linkInstance.Id);
                            }
                        }
                        
                        // 查找並添加連結的管線元件
                        var connectedElements = GetConnectedMEPElements(_clashInfo.MEPElement, document);
                        foreach (var connectedId in connectedElements)
                        {
                            var connectedElement = _clashInfo.MEPElement.Document.GetElement(connectedId);
                            if (connectedElement != null)
                            {
                                if (IsElementInCurrentDocument(connectedElement, document))
                                {
                                    elementIds.Add(connectedId);
                                }
                                else
                                {
                                    var linkInstance = GetLinkInstanceForElement(connectedElement, document);
                                    if (linkInstance != null && !linkInstanceIds.Contains(linkInstance.Id))
                                    {
                                        linkInstanceIds.Add(linkInstance.Id);
                                    }
                                }
                            }
                        }
                        
                        // 合併所有要選取的元件
                        elementIds.AddRange(linkInstanceIds);
                        
                        if (elementIds.Any())
                        {
                            uiDoc.Selection.SetElementIds(elementIds);
                        }

                        // 縮放到適合
                        uiDoc.GetOpenUIViews().FirstOrDefault()?.ZoomToFit();

                        _onCompleted?.Invoke("成功定位到碰撞點並選取相關元件");
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
                _onError?.Invoke($"定位到碰撞點時發生錯誤：{ex.Message}");
            }
        }

        public string GetName()
        {
            return "3D View Navigation Event";
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

        private List<ElementId> GetConnectedMEPElements(Element mepElement, Document document)
        {
            var connectedIds = new List<ElementId>();
            
            try
            {
                // 獲取MEP元件的連接器
                var connectorManager = GetConnectorManager(mepElement);
                if (connectorManager == null) return connectedIds;

                foreach (Connector connector in connectorManager.Connectors)
                {
                    // 獲取連接到此連接器的其他連接器
                    var connectedConnectors = connector.AllRefs;
                    foreach (Connector connectedConnector in connectedConnectors)
                    {
                        if (connectedConnector.Owner.Id != mepElement.Id)
                        {
                            // 檢查連接的元件是否為MEP元件
                            var connectedElement = connectedConnector.Owner;
                            if (IsMEPElement(connectedElement))
                            {
                                connectedIds.Add(connectedElement.Id);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting connected elements: {ex.Message}");
            }

            return connectedIds;
        }

        private ConnectorManager GetConnectorManager(Element element)
        {
            if (element is Pipe pipe) return pipe.ConnectorManager;
            if (element is Duct duct) return duct.ConnectorManager;
            if (element is CableTray cableTray) return cableTray.ConnectorManager;
            if (element is Conduit conduit) return conduit.ConnectorManager;
            if (element is FlexPipe flexPipe) return flexPipe.ConnectorManager;
            
            return null;
        }

        private bool IsMEPElement(Element element)
        {
            return element is Pipe || 
                   element is Duct || 
                   element is CableTray || 
                   element is Conduit || 
                   element is FlexPipe ||
                   element is FamilyInstance; // 包括MEP配件
        }

        private bool IsElementInCurrentDocument(Element element, Document currentDocument)
        {
            return element.Document.PathName == currentDocument.PathName;
        }

        private RevitLinkInstance GetLinkInstanceForElement(Element element, Document currentDocument)
        {
            try
            {
                var elementDocument = element.Document;
                
                // 查找對應的連結實例
                var linkCollector = new FilteredElementCollector(currentDocument)
                    .OfClass(typeof(RevitLinkInstance));

                foreach (RevitLinkInstance linkInstance in linkCollector)
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc != null && linkDoc.PathName == elementDocument.PathName)
                    {
                        return linkInstance;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error finding link instance: {ex.Message}");
            }

            return null;
        }
    }

    public class View3DNavigationHandler
    {
        private readonly View3DNavigationEvent _externalEvent;
        private readonly ExternalEvent _revitEvent;

        public View3DNavigationHandler()
        {
            _externalEvent = new View3DNavigationEvent();
            _revitEvent = ExternalEvent.Create(_externalEvent);
        }

        public async void NavigateToClash(ClashInfo clashInfo, Action<string> onCompleted, Action<string> onError)
        {
            _externalEvent.SetData(clashInfo, onCompleted, onError);
            
            // 使用 await 等待 External Event 完成
            var result = _revitEvent.Raise();
            if (result != ExternalEventRequest.Accepted)
            {
                onError?.Invoke("無法執行3D導航操作");
            }
        }
    }
}
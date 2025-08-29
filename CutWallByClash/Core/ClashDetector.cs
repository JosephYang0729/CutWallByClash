using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using CutWallByClash.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CutWallByClash.Core
{
    public class ClashDetector
    {
        private readonly Document _document;

        public ClashDetector(Document document)
        {
            _document = document;
        }

        public List<ClashInfo> DetectClashes(HashSet<MEPCategory> selectedCategories, IProgress<int> progress = null)
        {
            var clashes = new List<ClashInfo>();
            
            // 獲取所有牆（包括連結模型）
            var walls = GetAllWalls();
            
            // 獲取所有MEP元件
            var mepElements = GetMEPElements(selectedCategories);

            var totalCombinations = walls.Count * mepElements.Count;
            var processedCombinations = 0;

            foreach (var wall in walls)
            {
                foreach (var mepElement in mepElements)
                {
                    var clash = CheckClash(wall, mepElement);
                    if (clash != null)
                    {
                        clashes.Add(clash);
                    }
                    
                    processedCombinations++;
                    if (totalCombinations > 0)
                    {
                        var progressValue = (int)((double)processedCombinations / totalCombinations * 100);
                        progress?.Report(progressValue);
                    }
                }
            }

            return clashes;
        }

        private List<Wall> GetAllWalls()
        {
            var walls = new List<Wall>();
            
            // 獲取當前文檔的牆
            var wallCollector = new FilteredElementCollector(_document)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(w => w.WallType != null && (w.WallType.Kind == WallKind.Basic || w.WallType.Kind == WallKind.Curtain));
            
            walls.AddRange(wallCollector);

            // 獲取連結模型的牆
            var linkCollector = new FilteredElementCollector(_document)
                .OfClass(typeof(RevitLinkInstance));

            foreach (RevitLinkInstance linkInstance in linkCollector)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc != null)
                {
                    var linkWalls = new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(Wall))
                        .Cast<Wall>()
                        .Where(w => w.WallType != null && (w.WallType.Kind == WallKind.Basic || w.WallType.Kind == WallKind.Curtain));
                    
                    walls.AddRange(linkWalls);
                }
            }

            return walls;
        }

        private List<Element> GetMEPElements(HashSet<MEPCategory> selectedCategories)
        {
            var elements = new List<Element>();

            if (selectedCategories.Contains(MEPCategory.Pipes))
            {
                elements.AddRange(GetPipes());
            }
            if (selectedCategories.Contains(MEPCategory.Ducts))
            {
                elements.AddRange(GetDucts());
            }
            if (selectedCategories.Contains(MEPCategory.CableTray))
            {
                elements.AddRange(GetCableTrays());
            }
            if (selectedCategories.Contains(MEPCategory.Conduit))
            {
                elements.AddRange(GetConduits());
            }
            if (selectedCategories.Contains(MEPCategory.FlexPipe))
            {
                elements.AddRange(GetFlexPipes());
            }

            return elements;
        }

        private List<Element> GetPipes()
        {
            var pipes = new List<Element>();
            
            // 當前文檔的管
            pipes.AddRange(new FilteredElementCollector(_document)
                .OfClass(typeof(Pipe))
                .Cast<Element>());

            // 連結模型的管
            var linkCollector = new FilteredElementCollector(_document)
                .OfClass(typeof(RevitLinkInstance));

            foreach (RevitLinkInstance linkInstance in linkCollector)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc != null)
                {
                    pipes.AddRange(new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(Pipe))
                        .Cast<Element>());
                }
            }

            return pipes;
        }

        private List<Element> GetDucts()
        {
            var ducts = new List<Element>();
            
            ducts.AddRange(new FilteredElementCollector(_document)
                .OfClass(typeof(Duct))
                .Cast<Element>());

            var linkCollector = new FilteredElementCollector(_document)
                .OfClass(typeof(RevitLinkInstance));

            foreach (RevitLinkInstance linkInstance in linkCollector)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc != null)
                {
                    ducts.AddRange(new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(Duct))
                        .Cast<Element>());
                }
            }

            return ducts;
        }

        private List<Element> GetCableTrays()
        {
            var cableTrays = new List<Element>();
            
            cableTrays.AddRange(new FilteredElementCollector(_document)
                .OfClass(typeof(CableTray))
                .Cast<Element>());

            var linkCollector = new FilteredElementCollector(_document)
                .OfClass(typeof(RevitLinkInstance));

            foreach (RevitLinkInstance linkInstance in linkCollector)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc != null)
                {
                    cableTrays.AddRange(new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(CableTray))
                        .Cast<Element>());
                }
            }

            return cableTrays;
        }

        private List<Element> GetConduits()
        {
            var conduits = new List<Element>();
            
            conduits.AddRange(new FilteredElementCollector(_document)
                .OfClass(typeof(Conduit))
                .Cast<Element>());

            var linkCollector = new FilteredElementCollector(_document)
                .OfClass(typeof(RevitLinkInstance));

            foreach (RevitLinkInstance linkInstance in linkCollector)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc != null)
                {
                    conduits.AddRange(new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(Conduit))
                        .Cast<Element>());
                }
            }

            return conduits;
        }

        private List<Element> GetFlexPipes()
        {
            var flexPipes = new List<Element>();
            
            // 獲取撓性管 - 通常是 FlexPipe 類型
            flexPipes.AddRange(new FilteredElementCollector(_document)
                .OfClass(typeof(FlexPipe))
                .Cast<Element>());

            var linkCollector = new FilteredElementCollector(_document)
                .OfClass(typeof(RevitLinkInstance));

            foreach (RevitLinkInstance linkInstance in linkCollector)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc != null)
                {
                    flexPipes.AddRange(new FilteredElementCollector(linkDoc)
                        .OfClass(typeof(FlexPipe))
                        .Cast<Element>());
                }
            }

            return flexPipes;
        }

        private ClashInfo CheckClash(Wall wall, Element mepElement)
        {
            try
            {
                var wallSolid = GeometryUtils.GetElementSolid(wall);
                var mepSolid = GeometryUtils.GetElementSolid(mepElement);

                if (GeometryUtils.DoSolidsIntersect(wallSolid, mepSolid))
                {
                    var clashPoint = GeometryUtils.GetIntersectionCentroid(wallSolid, mepSolid);
                    var category = GetMEPCategory(mepElement);
                    var dimensions = GetMEPDimensions(mepElement);

                    return new ClashInfo
                    {
                        Wall = wall,
                        MEPElement = mepElement,
                        ClashPoint = clashPoint,
                        Category = category,
                        MEPBoundingBox = mepElement.get_BoundingBox(null) ?? new BoundingBoxXYZ(),
                        ElementDiameter = dimensions.diameter,
                        ElementWidth = dimensions.width,
                        ElementHeight = dimensions.height
                    };
                }
            }
            catch (Exception ex)
            {
                // 記錄錯誤但繼續處理其他元件
                System.Diagnostics.Debug.WriteLine($"Clash detection error for wall {wall.Id} and element {mepElement.Id}: {ex.Message}");
            }

            return null;
        }


        private MEPCategory GetMEPCategory(Element element)
        {
            if (element.GetType() == typeof(Pipe)) return MEPCategory.Pipes;
            if (element.GetType() == typeof(Duct)) return MEPCategory.Ducts;
            if (element.GetType() == typeof(CableTray)) return MEPCategory.CableTray;
            if (element.GetType() == typeof(Conduit)) return MEPCategory.Conduit;
            if (element.GetType() == typeof(FlexPipe)) return MEPCategory.FlexPipe;
            
            return MEPCategory.Pipes; // 預設值
        }

        private (double diameter, double width, double height) GetMEPDimensions(Element element)
        {
            double diameter = 0, width = 0, height = 0;

            var pipe = element as Pipe;
            if (pipe != null)
            {
                diameter = pipe.Diameter * 304.8; // 轉換為mm
            }
            else
            {
                var duct = element as Duct;
                if (duct != null)
                {
                    var connectors = duct.ConnectorManager?.Connectors;
                    Connector connector = null;
                    if (connectors != null)
                    {
                        foreach (Connector conn in connectors)
                        {
                            connector = conn;
                            break;
                        }
                    }
                    if (connector != null)
                    {
                        if (connector.Shape == ConnectorProfileType.Round)
                        {
                            diameter = connector.Radius * 2 * 304.8; // 轉換為mm
                        }
                        else
                        {
                            width = connector.Width * 304.8; // 轉換為mm
                            height = connector.Height * 304.8; // 轉換為mm
                        }
                    }
                }
                else
                {
                    var cableTray = element as CableTray;
                    if (cableTray != null)
                    {
                        width = cableTray.Width * 304.8; // 轉換為mm
                        height = cableTray.Height * 304.8; // 轉換為mm
                    }
                    else
                    {
                        var conduit = element as Conduit;
                        if (conduit != null)
                        {
                            diameter = conduit.Diameter * 304.8; // 轉換為mm
                        }
                        else
                        {
                            var flexPipe = element as FlexPipe;
                            if (flexPipe != null)
                            {
                                diameter = flexPipe.Diameter * 304.8; // 轉換為mm
                            }
                        }
                    }
                }
            }

            return (diameter, width, height);
        }
    }
}
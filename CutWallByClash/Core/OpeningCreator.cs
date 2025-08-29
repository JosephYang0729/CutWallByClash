using Autodesk.Revit.DB;
using CutWallByClash.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CutWallByClash.Core
{
    public enum OpeningResult
    {
        Success,
        Failed,
        Skipped
    }

    public class OpeningCreator
    {
        private readonly Document _document;

        public OpeningCreator(Document document)
        {
            _document = document;
        }

        public List<OpeningGroup> GroupClashes(List<ClashInfo> clashes, OpeningParameters parameters)
        {
            var groups = new List<OpeningGroup>();
            var processedClashes = new HashSet<ClashInfo>();

            foreach (var clash in clashes)
            {
                if (processedClashes.Contains(clash))
                    continue;

                // 根據MEP元件類型自動決定開口類型
                var openingType = DetermineOpeningType(clash);
                
                var group = new OpeningGroup
                {
                    Wall = clash.Wall,
                    OpeningType = openingType
                };

                group.Clashes.Add(clash);
                processedClashes.Add(clash);

                // 尋找附近的碰撞點
                var nearbyClashes = clashes.Where(c => 
                    c.Wall.Id == clash.Wall.Id && 
                    !processedClashes.Contains(c) &&
                    c.ClashPoint.DistanceTo(clash.ClashPoint) <= GeometryUtils.MmToFeet(parameters.MergeDistance))
                    .ToList();

                foreach (var nearbyClash in nearbyClashes)
                {
                    group.Clashes.Add(nearbyClash);
                    processedClashes.Add(nearbyClash);
                }

                // 計算群組的中心點和所需尺寸
                CalculateGroupParameters(group, parameters);
                groups.Add(group);
            }

            // 檢查圓形開口合併邏輯
            groups = CheckAndMergeCircularOpenings(groups, parameters);

            return groups;
        }

        /// <summary>
        /// 根據MEP元件類型自動決定開口類型
        /// </summary>
        private OpeningType DetermineOpeningType(ClashInfo clash)
        {
            switch (clash.Category)
            {
                case MEPCategory.Ducts:      // 風管 -> 矩形開口
                case MEPCategory.CableTray:  // 電纜架 -> 矩形開口
                    return OpeningType.Rectangular;
                
                case MEPCategory.Pipes:      // 管道 -> 圓形開口
                case MEPCategory.Conduit:    // 電管 -> 圓形開口
                case MEPCategory.FlexPipe:   // 撓性管 -> 圓形開口
                    return OpeningType.Circular;
                
                default:
                    return OpeningType.Circular; // 預設圓形
            }
        }

        /// <summary>
        /// 檢查並合併相近的圓形開口為矩形開口
        /// </summary>
        private List<OpeningGroup> CheckAndMergeCircularOpenings(List<OpeningGroup> groups, OpeningParameters parameters)
        {
            var mergedGroups = new List<OpeningGroup>();
            var processedGroups = new HashSet<OpeningGroup>();

            foreach (var group in groups)
            {
                if (processedGroups.Contains(group))
                    continue;

                if (group.OpeningType != OpeningType.Circular)
                {
                    mergedGroups.Add(group);
                    continue;
                }

                // 尋找相近的圓形開口
                var nearbyCircularGroups = groups.Where(g => 
                    g != group &&
                    !processedGroups.Contains(g) &&
                    g.OpeningType == OpeningType.Circular &&
                    g.Wall.Id == group.Wall.Id &&
                    ShouldMergeCircularOpenings(group, g))
                    .ToList();

                if (nearbyCircularGroups.Any())
                {
                    // 合併為矩形開口
                    var mergedGroup = MergeCircularGroupsToRectangular(group, nearbyCircularGroups, parameters);
                    mergedGroups.Add(mergedGroup);
                    
                    processedGroups.Add(group);
                    foreach (var nearbyGroup in nearbyCircularGroups)
                    {
                        processedGroups.Add(nearbyGroup);
                    }
                }
                else
                {
                    mergedGroups.Add(group);
                    processedGroups.Add(group);
                }
            }

            return mergedGroups;
        }

        /// <summary>
        /// 判斷兩個圓形開口是否應該合併 (r1 + r2 + D1 < 100mm)
        /// 其中 r1, r2 是管線半徑，D1 是兩個管心的距離
        /// </summary>
        private bool ShouldMergeCircularOpenings(OpeningGroup group1, OpeningGroup group2)
        {
            var distance = group1.CenterPoint.DistanceTo(group2.CenterPoint) * 304.8; // 轉換為mm (D1)
            
            // 獲取管線半徑（不是開口半徑）
            var r1 = group1.Clashes.First().ElementDiameter / 2; // 管線半徑1
            var r2 = group2.Clashes.First().ElementDiameter / 2; // 管線半徑2
            
            var totalDistance = r1 + r2 + distance;
            
            // 調試輸出
            System.Diagnostics.Debug.WriteLine($"圓形開口合併檢查：r1={r1}mm, r2={r2}mm, D1={distance}mm, 總距離={totalDistance}mm");
            
            return totalDistance < 100; // 小於100mm(10cm)就合併
        }

        /// <summary>
        /// 將多個圓形開口群組合併為一個矩形開口群組
        /// 計算能夠涵蓋所有圓形管線的最小矩形開口
        /// </summary>
        private OpeningGroup MergeCircularGroupsToRectangular(OpeningGroup mainGroup, List<OpeningGroup> nearbyGroups, OpeningParameters parameters)
        {
            var allGroups = new List<OpeningGroup> { mainGroup };
            allGroups.AddRange(nearbyGroups);

            // 計算所有碰撞點的邊界，考慮管線半徑
            var allClashes = allGroups.SelectMany(g => g.Clashes).ToList();
            
            // 計算每個管線的邊界（中心點 ± 半徑）
            var minX = double.MaxValue;
            var maxX = double.MinValue;
            var minY = double.MaxValue;
            var maxY = double.MinValue;
            var avgZ = allClashes.Average(c => c.ClashPoint.Z);

            foreach (var clash in allClashes)
            {
                var radius = GeometryUtils.MmToFeet(clash.ElementDiameter / 2); // 轉換為英尺
                
                minX = Math.Min(minX, clash.ClashPoint.X - radius);
                maxX = Math.Max(maxX, clash.ClashPoint.X + radius);
                minY = Math.Min(minY, clash.ClashPoint.Y - radius);
                maxY = Math.Max(maxY, clash.ClashPoint.Y + radius);
            }

            // 計算合併後的中心點
            var centerX = (minX + maxX) / 2;
            var centerY = (minY + maxY) / 2;
            var mergedCenter = new XYZ(centerX, centerY, avgZ);

            // 計算所需的矩形尺寸（包含所有管線的邊界 + 餘量）
            var spanX = (maxX - minX) * 304.8; // 轉換為mm
            var spanY = (maxY - minY) * 304.8;

            var requiredWidth = spanX + (parameters.RectangularWidth * 2);
            var requiredHeight = spanY + (parameters.RectangularHeight * 2);

            // 調試輸出
            System.Diagnostics.Debug.WriteLine($"合併矩形開口：涵蓋範圍 {spanX:F1}x{spanY:F1}mm，最終尺寸 {requiredWidth:F1}x{requiredHeight:F1}mm");

            return new OpeningGroup
            {
                Wall = mainGroup.Wall,
                OpeningType = OpeningType.Rectangular,
                CenterPoint = mergedCenter,
                Clashes = allClashes,
                RequiredWidth = requiredWidth,
                RequiredHeight = requiredHeight
            };
        }

        private void CalculateGroupParameters(OpeningGroup group, OpeningParameters parameters)
        {
            // 使用第一個碰撞點作為放置點位（不使用平均值）
            var firstClash = group.Clashes.First();
            group.CenterPoint = firstClash.ClashPoint;

            if (group.OpeningType == OpeningType.Rectangular)
            {
                // 對於矩形開口，使用MEP元件的實際尺寸
                if (firstClash.ElementDiameter > 0)
                {
                    // 圓形元件（管道、電管等）
                    var diameter = firstClash.ElementDiameter;
                    group.RequiredWidth = diameter + (parameters.RectangularWidth * 2);
                    group.RequiredHeight = diameter + (parameters.RectangularHeight * 2);
                }
                else
                {
                    // 矩形元件（風管、電纜架等）
                    group.RequiredWidth = firstClash.ElementWidth + (parameters.RectangularWidth * 2);
                    group.RequiredHeight = firstClash.ElementHeight + (parameters.RectangularHeight * 2);
                }
            }
            else
            {
                // 圓形開口直徑 = 管線直徑 + 開口距離 * 2
                var pipeDiameter = firstClash.ElementDiameter; // 直接使用碰撞點的管線直徑
                group.RequiredDiameter = pipeDiameter + (parameters.CircularDiameter * 2);
                
                // 調試輸出
                System.Diagnostics.Debug.WriteLine($"圓形開口計算：管線直徑={pipeDiameter}mm, 開口距離={parameters.CircularDiameter}mm, 最終開口直徑={group.RequiredDiameter}mm");
            }
        }

        private double GetElementActualWidth(ClashInfo clash)
        {
            if (clash.ElementDiameter > 0)
                return clash.ElementDiameter; // 圓形元件的寬度就是直徑
            return clash.ElementWidth; // 矩形元件的實際寬度
        }

        private double GetElementActualHeight(ClashInfo clash)
        {
            if (clash.ElementDiameter > 0)
                return clash.ElementDiameter; // 圓形元件的高度就是直徑
            return clash.ElementHeight; // 矩形元件的實際高度
        }

        private double GetElementActualDiameter(ClashInfo clash)
        {
            if (clash.ElementDiameter > 0)
                return clash.ElementDiameter; // 圓形元件的直徑
            // 矩形元件取較大的尺寸作為等效直徑
            return Math.Max(clash.ElementWidth, clash.ElementHeight);
        }

        private double GetElementHalfWidth(ClashInfo clash)
        {
            if (clash.ElementDiameter > 0)
                return GeometryUtils.MmToFeet(clash.ElementDiameter / 2);
            return GeometryUtils.MmToFeet(clash.ElementWidth / 2);
        }

        private double GetElementHalfHeight(ClashInfo clash)
        {
            if (clash.ElementDiameter > 0)
                return GeometryUtils.MmToFeet(clash.ElementDiameter / 2);
            return GeometryUtils.MmToFeet(clash.ElementHeight / 2);
        }

        public Dictionary<ClashInfo, ElementId> CreateOpenings(List<OpeningGroup> groups, IProgress<int> progress)
        {
            var openingIds = new Dictionary<ClashInfo, ElementId>();
            var totalGroups = groups.Count;
            var processedGroups = 0;
            var successCount = 0;
            var failedCount = 0;
            var skippedCount = 0;
            var mergedCount = 0; // 記錄合併的開口數量

            try
            {
                foreach (var group in groups)
                {
                    var (result, openingId) = CreateOpening(group);
                    switch (result)
                    {
                        case OpeningResult.Success:
                            successCount++;
                            // 為該群組的所有碰撞點記錄開口ID
                            if (openingId != null)
                            {
                                foreach (var clash in group.Clashes)
                                {
                                    openingIds[clash] = openingId;
                                }
                                
                                // 如果這是合併的群組（包含多個原始碰撞點），記錄合併資訊
                                if (group.Clashes.Count > 1)
                                {
                                    mergedCount++;
                                    System.Diagnostics.Debug.WriteLine($"合併開口：將 {group.Clashes.Count} 個碰撞點合併為一個{(group.OpeningType == OpeningType.Rectangular ? "矩形" : "圓形")}開口 (ID: {openingId})");
                                }
                            }
                            break;
                        case OpeningResult.Failed:
                            failedCount++;
                            break;
                        case OpeningResult.Skipped:
                            skippedCount++;
                            break;
                    }
                    
                    processedGroups++;
                    progress?.Report((int)((double)processedGroups / totalGroups * 100));
                }
                
                if (successCount == 0 && failedCount > 0)
                {
                    throw new Exception($"無法建立任何開口。成功: {successCount}, 失敗: {failedCount}, 跳過(連結模型): {skippedCount}");
                }
                
                // 在成功訊息中包含合併資訊
                if (mergedCount > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"開口建立完成！成功: {successCount}, 失敗: {failedCount}, 跳過: {skippedCount}, 智能合併: {mergedCount}");
                }
                
                return openingIds;
            }
            catch (Exception ex)
            {
                throw new Exception($"建立開口失敗: {ex.Message}。成功: {successCount}, 失敗: {failedCount}, 跳過(連結模型): {skippedCount}", ex);
            }
        }

        private (OpeningResult, ElementId) CreateOpening(OpeningGroup group)
        {
            try
            {
                var wall = group.Wall;
                var centerPoint = group.CenterPoint;

                // 檢查是否為連結模型的牆
                if (IsLinkedWall(wall))
                {
                    System.Diagnostics.Debug.WriteLine($"跳過連結模型的牆 {wall.Id}，無法在連結模型中建立開口");
                    return (OpeningResult.Skipped, null);
                }

                // 獲取牆的基本資訊
                var wallCurve = (wall.Location as LocationCurve)?.Curve;
                if (wallCurve == null) 
                {
                    System.Diagnostics.Debug.WriteLine($"無法獲取牆 {wall.Id} 的位置曲線");
                    return (OpeningResult.Failed, null);
                }

                // 計算牆面上的開口位置
                var openingCenter = CalculateOpeningCenterOnWall(wall, centerPoint);
                if (openingCenter == null)
                {
                    System.Diagnostics.Debug.WriteLine($"無法計算牆 {wall.Id} 上的開口位置");
                    return (OpeningResult.Failed, null);
                }
                // 獲取牆的方向向量
                var wallDirection = (wallCurve.GetEndPoint(1) - wallCurve.GetEndPoint(0)).Normalize();
                
                ElementId openingId;
                if (group.OpeningType == OpeningType.Rectangular)
                {
                    openingId = CreateRectangularOpening(wall, openingCenter, wallDirection, 
                        GeometryUtils.MmToFeet(group.RequiredWidth), GeometryUtils.MmToFeet(group.RequiredHeight));
                }
                else
                {
                    openingId = CreateCircularOpening(wall, openingCenter, 
                        GeometryUtils.MmToFeet(group.RequiredDiameter));
                }

                return (OpeningResult.Success, openingId);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"建立開口失敗: {ex.Message} - {ex.StackTrace}");
                return (OpeningResult.Failed, null);
            }
        }

        private XYZ CalculateOpeningCenterOnWall(Wall wall, XYZ clashPoint)
        {
            try
            {
                // 獲取牆的位置曲線
                var wallCurve = (wall.Location as LocationCurve)?.Curve;
                if (wallCurve == null) return null;

                // 找到牆曲線上最接近碰撞點的點
                var closestPoint = wallCurve.Project(clashPoint).XYZPoint;
                
                // 使用碰撞點的Z座標（高度）
                return new XYZ(closestPoint.X, closestPoint.Y, clashPoint.Z);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"計算開口中心點失敗: {ex.Message}");
                return null;
            }
        }

        private bool IsLinkedWall(Wall wall)
        {
            // 檢查牆是否來自連結模型
            return wall.Document.PathName != _document.PathName;
        }

        private ElementId CreateRectangularOpening(Wall wall, XYZ centerPoint, XYZ wallDirection, 
            double width, double height)
        {
            string ExportFolder = $"{Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)}\\Autodesk\\Revit\\Addins\\CutWallByClash\\";
            try
            {
                FilteredElementCollector famCollector = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilySymbol));
                FamilySymbol openingSymbol = famCollector
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(fs => fs.Name == "矩形開口");

                if (openingSymbol == null)
                {
                    Family framing_family = null;
                    IFamilyLoadOptions famLoadOptions = new FamilyOption();
                    string load_path = ExportFolder + "矩形開口.rfa";
                    _document.LoadFamily(load_path, famLoadOptions, out framing_family);
                    openingSymbol = _document.GetElement(framing_family.GetFamilySymbolIds().First()) as FamilySymbol;
                }

                if (!openingSymbol.IsActive)
                    openingSymbol.Activate();

                FamilyInstance openingInstance = _document.Create.NewFamilyInstance(
                    centerPoint, openingSymbol, wall, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                openingInstance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(centerPoint.Z - (height / 2));
                openingInstance.LookupParameter("開口寬度").Set(width);
                openingInstance.LookupParameter("開口長度").Set(height);

                _document.Regenerate();
                InstanceVoidCutUtils.AddInstanceVoidCut(_document, wall, openingInstance);
                
                return openingInstance.Id;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"建立矩形開口失敗: {ex.Message}");
                throw new Exception($"建立矩形開口失敗: {ex.Message}", ex);
            }
        }

        private ElementId CreateCircularOpening(Wall wall, XYZ centerPoint, double diameter)
        {
            string ExportFolder = $"{Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)}\\Autodesk\\Revit\\Addins\\CutWallByClash\\";
            try
            {
                FilteredElementCollector famCollector = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilySymbol));
                FamilySymbol openingSymbol = famCollector
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(fs => fs.Name == "圓形開口");

                if (openingSymbol == null)
                {
                    Family framing_family = null;
                    IFamilyLoadOptions famLoadOptions = new FamilyOption();
                    string load_path = ExportFolder + "圓形開口.rfa";
                    _document.LoadFamily(load_path, famLoadOptions, out framing_family);
                    openingSymbol = _document.GetElement(framing_family.GetFamilySymbolIds().First()) as FamilySymbol;
                }

                if (!openingSymbol.IsActive)
                    openingSymbol.Activate();

                FamilyInstance openingInstance = _document.Create.NewFamilyInstance(
                    centerPoint, openingSymbol, wall, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                openingInstance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(centerPoint.Z - (diameter / 2));
                openingInstance.LookupParameter("直徑").Set(diameter);

                _document.Regenerate();

                /*var walls = new List<Wall>();

                var wallCollector = new FilteredElementCollector(_document)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .Where(w => w.WallType != null && (w.WallType.Kind == WallKind.Basic || w.WallType.Kind == WallKind.Curtain));

                walls.AddRange(wallCollector);

                foreach (Wall model_wall in walls)
                {
                    try
                    {
                        InstanceVoidCutUtils.AddInstanceVoidCut(_document, model_wall, openingInstance);
                    }
                    catch { }
                }*/
                InstanceVoidCutUtils.AddInstanceVoidCut(_document, wall, openingInstance);
                
                return openingInstance.Id;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"建立圓形開口失敗: {ex.Message}");
                throw new Exception($"建立圓形開口失敗: {ex.Message}", ex);
            }
        }

        private class FamilyOption : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true;
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family;
                overwriteParameterValues = true;
                return true;
            }
        }
    }
}
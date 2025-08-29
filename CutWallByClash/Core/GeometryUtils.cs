using Autodesk.Revit.DB;
using System;

namespace CutWallByClash.Core
{
    public static class GeometryUtils
    {
        /// <summary>
        /// 英尺轉換為毫米
        /// </summary>
        public static double FeetToMm(double feet)
        {
            return feet * 304.8;
        }

        /// <summary>
        /// 毫米轉換為英尺
        /// </summary>
        public static double MmToFeet(double mm)
        {
            return mm / 304.8;
        }

        /// <summary>
        /// 安全地獲取元件的幾何體
        /// </summary>
        public static Solid GetElementSolid(Element element)
        {
            if (element == null) return null;

            var options = new Options
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            try
            {
                var geometry = element.get_Geometry(options);
                if (geometry == null) return null;

                foreach (GeometryObject geoObj in geometry)
                {
                    var solid = geoObj as Solid;
                    if (solid != null && solid.Volume > 0.001)
                    {
                        return solid;
                    }
                    else
                    {
                        var instance = geoObj as GeometryInstance;
                        if (instance != null)
                        {
                            var instanceGeometry = instance.GetInstanceGeometry();
                            if (instanceGeometry != null)
                            {
                                foreach (GeometryObject instObj in instanceGeometry)
                                {
                                    var instSolid = instObj as Solid;
                                    if (instSolid != null && instSolid.Volume > 0.001)
                                    {
                                        return instSolid;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting geometry for element {element.Id}: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// 檢查兩個實體是否相交
        /// </summary>
        public static bool DoSolidsIntersect(Solid solid1, Solid solid2)
        {
            if (solid1 == null || solid2 == null) return false;

            try
            {
                var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                    solid1, solid2, BooleanOperationsType.Intersect);
                
                return intersection != null && intersection.Volume > 0.001;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking solid intersection: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 獲取兩個實體的交集中心點
        /// </summary>
        public static XYZ GetIntersectionCentroid(Solid solid1, Solid solid2)
        {
            try
            {
                var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                    solid1, solid2, BooleanOperationsType.Intersect);
                
                if (intersection != null && intersection.Volume > 0.001)
                {
                    return intersection.ComputeCentroid();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting intersection centroid: {ex.Message}");
            }

            // 如果無法計算交集，返回第二個實體的中心點
            return solid2?.ComputeCentroid() ?? XYZ.Zero;
        }
    }
}
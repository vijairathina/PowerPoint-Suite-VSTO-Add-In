using System;
using System.Collections.Generic;
using OneAssist.SlideCompare.Matching;
using OneAssist.SlideCompare.Models;
using OneAssist.SlideCompare.Services;

namespace OneAssist.SlideCompare.Comparison
{
    /// <summary>
    /// Coordinates shape-level comparisons between shape collections.
    /// </summary>
    public static class ShapeComparisonEngine
    {
        public static List<ShapeComparisonResult> CompareShapeCollections(
            List<ShapeSnapshot> sourceShapes,
            List<ShapeSnapshot> destShapes,
            int sourceSlideNumber,
            int destSlideNumber,
            CompareSettings settings)
        {
            return ShapeMatchingEngine.MatchShapes(sourceShapes, destShapes, sourceSlideNumber, destSlideNumber, settings);
        }
    }
}

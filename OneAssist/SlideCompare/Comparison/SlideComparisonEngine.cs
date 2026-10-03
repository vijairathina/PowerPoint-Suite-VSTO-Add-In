using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.SlideCompare.Matching;
using OneAssist.SlideCompare.Models;
using OneAssist.SlideCompare.Services;

namespace OneAssist.SlideCompare.Comparison
{
    /// <summary>
    /// Coordinates comparison between two slides once paired.
    /// </summary>
    public static class SlideComparisonEngine
    {
        public static SlideComparisonResult CompareSlides(
            SlideSnapshot source,
            SlideSnapshot dest,
            CompareSettings settings)
        {
            settings = settings ?? new CompareSettings();

            if (source == null && dest == null) return new SlideComparisonResult();

            if (source == null)
            {
                // Added
                return new SlideComparisonResult
                {
                    SourceSlideNumber = 0,
                    DestinationSlideNumbers = new List<int> { dest.SlideNumber },
                    MatchType = SlideMatchType.Added,
                    DestinationTitle = dest.Title,
                    SummaryText = $"Slide {dest.SlideNumber} added in destination."
                };
            }

            if (dest == null)
            {
                // Deleted
                return new SlideComparisonResult
                {
                    SourceSlideNumber = source.SlideNumber,
                    DestinationSlideNumbers = new List<int>(),
                    MatchType = SlideMatchType.Deleted,
                    SourceTitle = source.Title,
                    SummaryText = $"Slide {source.SlideNumber} deleted in destination."
                };
            }

            double score = SlideMappingEngine.CalculateSlideSimilarity(source, dest, settings);
            var shapeResults = ShapeMatchingEngine.MatchShapes(
                source.GetAllShapesFlattened(),
                dest.GetAllShapesFlattened(),
                source.SlideNumber,
                dest.SlideNumber,
                settings);

            bool isReordered = (source.SlideNumber != dest.SlideNumber);
            bool hasIssues = shapeResults.Any(s => s.IsIssue);

            SlideMatchType matchType;
            if (!hasIssues)
            {
                matchType = isReordered ? SlideMatchType.Reordered : SlideMatchType.Matched;
            }
            else
            {
                matchType = isReordered ? SlideMatchType.Reordered : SlideMatchType.Modified;
            }

            return new SlideComparisonResult
            {
                SourceSlideNumber = source.SlideNumber,
                SourceSlideNumbers = new List<int> { source.SlideNumber },
                DestinationSlideNumbers = new List<int> { dest.SlideNumber },
                MatchType = matchType,
                SimilarityScore = score,
                SourceTitle = source.Title,
                DestinationTitle = dest.Title,
                ShapeResults = shapeResults,
                SummaryText = isReordered ? $"Slide moved from {source.SlideNumber} to {dest.SlideNumber}." : (hasIssues ? "Slide modified." : "Slide matched.")
            };
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OneAssist.SlideCompare.Matching;
using OneAssist.SlideCompare.Models;
using OneAssist.SlideCompare.Services;

namespace OneAssist.SlideCompare.Comparison
{
    /// <summary>
    /// Master comparison engine coordinating slide mapping, shape comparisons, and metrics aggregation.
    /// </summary>
    public static class ComparisonEngine
    {
        public static PresentationComparisonResult ComparePresentations(
            PresentationSnapshot source,
            PresentationSnapshot destination,
            CompareSettings settings,
            IProgress<Tuple<int, string>> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            settings = settings ?? new CompareSettings();

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(Tuple.Create(10, "Analyzing presentation structure..."));

            var result = new PresentationComparisonResult
            {
                SourcePresentationName = source.PresentationName,
                DestinationPresentationName = destination.PresentationName,
                ComparedAt = DateTime.UtcNow,
                TotalSourceSlides = source.Slides?.Count ?? 0,
                TotalDestinationSlides = destination.Slides?.Count ?? 0
            };

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(Tuple.Create(30, "Matching slides and analyzing splits/merges..."));

            var slideResults = SlideMappingEngine.MapSlides(source, destination, settings);
            result.SlideResults = slideResults;

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(Tuple.Create(80, "Aggregating metrics and differences..."));

            // Calculate aggregate counters
            int matchedContent = slideResults.Sum(s => s.ShapeResults?.Count(sh => sh.IsMatched) ?? 0);
            int totalContent = slideResults.Sum(s => s.ShapeResults?.Count ?? 0);
            int matchedSlides = slideResults.Count(s => s.MatchType == SlideMatchType.Matched);

            result.MatchedSlidesCount = matchedSlides;
            result.MatchedContentCount = matchedContent;
            result.TotalContentCount = totalContent;
            result.MatchedCount = matchedSlides;

            int deletedSlides = slideResults.Count(s => s.MatchType == SlideMatchType.Deleted);
            int deletedShapes = slideResults.Sum(s => s.ShapeResults?.Count(sh => sh.MatchType == ShapeMatchType.Deleted || sh.IsDeleted) ?? 0);
            int textDeletions = slideResults.Sum(s => s.ShapeResults?.Count(sh => sh.MatchType != ShapeMatchType.Deleted && sh.DiffTokens != null && sh.DiffTokens.Any(t => t.DiffType == TextDiffType.Deleted)) ?? 0);
            result.DeletedCount = deletedSlides + deletedShapes + textDeletions;

            int addedSlides = slideResults.Count(s => s.MatchType == SlideMatchType.Added);
            int addedShapes = slideResults.Sum(s => s.ShapeResults?.Count(sh => sh.MatchType == ShapeMatchType.Added || sh.IsAdded) ?? 0);
            int textAdditions = slideResults.Sum(s => s.ShapeResults?.Count(sh => sh.MatchType != ShapeMatchType.Added && sh.DiffTokens != null && sh.DiffTokens.Any(t => t.DiffType == TextDiffType.Added)) ?? 0);
            result.AddedCount = addedSlides + addedShapes + textAdditions;

            result.ModifiedCount = slideResults.Count(s => s.MatchType == SlideMatchType.Modified);
            result.MovedCount = slideResults.Count(s => s.MatchType == SlideMatchType.Moved || (s.ShapeResults != null && s.ShapeResults.Any(sh => sh.MatchType == ShapeMatchType.Moved)));
            result.ReorderedCount = slideResults.Count(s => s.MatchType == SlideMatchType.Reordered);
            result.SplitCount = slideResults.Count(s => s.MatchType == SlideMatchType.Split);
            result.MergedCount = slideResults.Count(s => s.MatchType == SlideMatchType.Merged);

            progress?.Report(Tuple.Create(100, "Comparison complete."));
            return result;
        }
    }
}

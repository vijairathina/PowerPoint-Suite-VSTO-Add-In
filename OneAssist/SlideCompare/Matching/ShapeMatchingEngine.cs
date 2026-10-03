using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.SlideCompare.Comparison;
using OneAssist.SlideCompare.Models;
using OneAssist.SlideCompare.Services;

namespace OneAssist.SlideCompare.Matching
{
    /// <summary>
    /// Matches shapes between source and destination slides using multi-dimensional weighted scoring,
    /// prioritizing best shape-to-shape pairing and comprehensive difference detection.
    /// </summary>
    public static class ShapeMatchingEngine
    {
        public static List<ShapeComparisonResult> MatchShapes(
            List<ShapeSnapshot> sourceShapes,
            List<ShapeSnapshot> destShapes,
            int sourceSlideNumber,
            int destSlideNumber,
            CompareSettings settings)
        {
            var results = new List<ShapeComparisonResult>();
            sourceShapes = sourceShapes ?? new List<ShapeSnapshot>();
            destShapes = destShapes ?? new List<ShapeSnapshot>();
            settings = settings ?? new CompareSettings();

            var matchedPairs = new List<Tuple<ShapeSnapshot, ShapeSnapshot, double>>();
            var claimedSource = new HashSet<ShapeSnapshot>();
            var claimedDest = new HashSet<ShapeSnapshot>();

            // 1. Calculate similarity matrix between all source and destination shapes
            var candidatePairs = new List<Tuple<ShapeSnapshot, ShapeSnapshot, double>>();
            foreach (var src in sourceShapes)
            {
                foreach (var dst in destShapes)
                {
                    double sim = CalculateShapeSimilarity(src, dst, settings);
                    candidatePairs.Add(Tuple.Create(src, dst, sim));
                }
            }

            // Sort candidate pairs: highest similarity first (global best matching)
            candidatePairs.Sort((a, b) => b.Item3.CompareTo(a.Item3));

            // Pass 1: Accept strong matches
            foreach (var pair in candidatePairs)
            {
                var src = pair.Item1;
                var dst = pair.Item2;
                double score = pair.Item3;

                if (claimedSource.Contains(src) || claimedDest.Contains(dst))
                    continue;

                bool exactText = src.HasText && dst.HasText &&
                                 string.Equals(src.NormalizedText, dst.NormalizedText, StringComparison.OrdinalIgnoreCase);

                if (score >= 0.40 || exactText)
                {
                    claimedSource.Add(src);
                    claimedDest.Add(dst);
                    matchedPairs.Add(pair);
                }
            }

            // Pass 2: Secondary pairing for remaining shapes on the slide
            // Pair remaining shapes based on placeholder role, shape type, position, or lone remaining text shapes
            var remainingSrc = sourceShapes.Where(s => !claimedSource.Contains(s)).ToList();
            var remainingDst = destShapes.Where(d => !claimedDest.Contains(d)).ToList();

            var secondaryPairs = new List<Tuple<ShapeSnapshot, ShapeSnapshot, double>>();
            foreach (var src in remainingSrc)
            {
                foreach (var dst in remainingDst)
                {
                    double secScore = CalculateShapeSimilarity(src, dst, settings);

                    bool samePlaceholder = !string.IsNullOrEmpty(src.PlaceholderType) &&
                                           string.Equals(src.PlaceholderType, dst.PlaceholderType, StringComparison.OrdinalIgnoreCase);
                    bool sameType = string.Equals(src.ShapeType, dst.ShapeType, StringComparison.OrdinalIgnoreCase);

                    if (samePlaceholder) secScore += 0.40;
                    else if (sameType) secScore += 0.20;

                    // If exactly 1 shape with text remains on both sides, pair them so the modified content is shown!
                    if (remainingSrc.Count == 1 && remainingDst.Count == 1 && (src.HasText || dst.HasText))
                    {
                        secScore += 0.35;
                    }

                    if (secScore >= 0.25)
                    {
                        secondaryPairs.Add(Tuple.Create(src, dst, Math.Min(1.0, secScore)));
                    }
                }
            }

            secondaryPairs.Sort((a, b) => b.Item3.CompareTo(a.Item3));

            foreach (var pair in secondaryPairs)
            {
                if (claimedSource.Contains(pair.Item1) || claimedDest.Contains(pair.Item2))
                    continue;

                claimedSource.Add(pair.Item1);
                claimedDest.Add(pair.Item2);
                matchedPairs.Add(pair);
            }

            // 2. Process matched pairs and generate detailed comparison results
            foreach (var pair in matchedPairs)
            {
                var src = pair.Item1;
                var dst = pair.Item2;
                double score = pair.Item3;

                var shapeRes = CompareMatchedShapes(src, dst, sourceSlideNumber, destSlideNumber, score, settings);
                results.Add(shapeRes);
            }

            // 3. Unmatched Source shapes -> DELETED
            foreach (var src in sourceShapes)
            {
                if (!claimedSource.Contains(src))
                {
                    var deletedRes = new ShapeComparisonResult
                    {
                        SourceShapeId = src.Id,
                        SourceShapeName = src.Name,
                        DestinationShapeId = 0,
                        DestinationShapeName = string.Empty,
                        MatchType = ShapeMatchType.Deleted,
                        DifferenceCategory = DifferenceCategory.Structure,
                        SimilarityScore = 0.0,
                        SourceText = src.RawText,
                        DestinationText = string.Empty,
                        TextDiff = !string.IsNullOrEmpty(src.RawText) ? $"- {src.RawText.Trim()}" : string.Empty,
                        DiffTokens = !string.IsNullOrEmpty(src.RawText)
                            ? new List<TextDiffToken> { new TextDiffToken { DiffType = TextDiffType.Deleted, Text = src.RawText } }
                            : new List<TextDiffToken>(),
                        SourceSlideNumber = sourceSlideNumber,
                        DestinationSlideNumber = destSlideNumber,
                        DestinationShapeIndex = 0,
                        Details = $"Shape '{src.Name}' ({src.ShapeType}) removed from slide."
                    };
                    results.Add(deletedRes);
                }
            }

            // 4. Unmatched Destination shapes -> ADDED
            foreach (var dst in destShapes)
            {
                if (!claimedDest.Contains(dst))
                {
                    var addedRes = new ShapeComparisonResult
                    {
                        SourceShapeId = 0,
                        SourceShapeName = string.Empty,
                        DestinationShapeId = dst.Id,
                        DestinationShapeName = dst.Name,
                        MatchType = ShapeMatchType.Added,
                        DifferenceCategory = DifferenceCategory.Structure,
                        SimilarityScore = 0.0,
                        SourceText = string.Empty,
                        DestinationText = dst.RawText,
                        TextDiff = !string.IsNullOrEmpty(dst.RawText) ? $"+ {dst.RawText.Trim()}" : string.Empty,
                        DiffTokens = !string.IsNullOrEmpty(dst.RawText)
                            ? new List<TextDiffToken> { new TextDiffToken { DiffType = TextDiffType.Added, Text = dst.RawText } }
                            : new List<TextDiffToken>(),
                        SourceSlideNumber = sourceSlideNumber,
                        DestinationSlideNumber = destSlideNumber,
                        DestinationShapeIndex = dst.Index,
                        Details = $"Shape '{dst.Name}' ({dst.ShapeType}) added to slide."
                    };
                    results.Add(addedRes);
                }
            }

            return results;
        }

        public static double CalculateShapeSimilarity(ShapeSnapshot src, ShapeSnapshot dst, CompareSettings settings)
        {
            if (src == null || dst == null) return 0.0;

            // Dimension 1: Text Similarity (40%)
            double textSim = 1.0;
            if (src.HasText || dst.HasText)
            {
                if (src.HasText && dst.HasText)
                {
                    if (string.Equals(src.NormalizedText, dst.NormalizedText, StringComparison.OrdinalIgnoreCase))
                    {
                        textSim = 1.0;
                    }
                    else
                    {
                        double lev = SimilarityCalculator.CalculateLevenshteinSimilarity(src.NormalizedText, dst.NormalizedText);
                        double dice = SimilarityCalculator.CalculateTextDiceSimilarity(src.NormalizedText, dst.NormalizedText);
                        textSim = Math.Max(lev, dice);
                    }
                }
                else
                {
                    textSim = 0.0;
                }
            }

            // Dimension 2: Shape Type Similarity (15%)
            double typeSim = string.Equals(src.ShapeType, dst.ShapeType, StringComparison.OrdinalIgnoreCase) ? 1.0 : 0.0;

            // Dimension 3: Position Similarity (15%)
            float dx = Math.Abs(src.Left - dst.Left);
            float dy = Math.Abs(src.Top - dst.Top);
            double posSim = Math.Max(0.0, 1.0 - ((dx / 960.0 * 0.5) + (dy / 540.0 * 0.5)));

            // Dimension 4: Size Similarity (10%)
            float dw = Math.Abs(src.Width - dst.Width);
            float dh = Math.Abs(src.Height - dst.Height);
            double sizeSim = Math.Max(0.0, 1.0 - ((dw / 960.0 * 0.5) + (dh / 540.0 * 0.5)));

            // Dimension 5: Name / Placeholder Similarity (10%)
            double nameSim = 0.0;
            if (!string.IsNullOrEmpty(src.PlaceholderType) && !string.IsNullOrEmpty(dst.PlaceholderType))
            {
                nameSim = string.Equals(src.PlaceholderType, dst.PlaceholderType, StringComparison.OrdinalIgnoreCase) ? 1.0 : 0.0;
            }
            else
            {
                nameSim = SimilarityCalculator.CalculateTextDiceSimilarity(src.Name, dst.Name);
            }

            // Dimension 6: Visual / Content Similarity (10%)
            double visualSim = 0.5;
            if (src.IsTable && dst.IsTable)
            {
                visualSim = (src.TableRowCount == dst.TableRowCount && src.TableColumnCount == dst.TableColumnCount) ? 1.0 : 0.6;
            }
            else if (src.IsChart && dst.IsChart)
            {
                visualSim = string.Equals(src.ChartTitle, dst.ChartTitle, StringComparison.OrdinalIgnoreCase) ? 1.0 : 0.7;
            }
            else if (src.IsImage && dst.IsImage)
            {
                bool hashMatch = !string.IsNullOrEmpty(src.ImageHash) && string.Equals(src.ImageHash, dst.ImageHash, StringComparison.OrdinalIgnoreCase);
                visualSim = hashMatch ? 1.0 : 0.3;
                if (!src.HasText && !dst.HasText)
                {
                    textSim = visualSim;
                }
            }

            double totalScore =
                (textSim * settings.WeightText) +
                (typeSim * settings.WeightShapeType) +
                (posSim * settings.WeightPosition) +
                (sizeSim * settings.WeightSize) +
                (nameSim * settings.WeightNamePlaceholder) +
                (visualSim * settings.WeightVisualContent);

            return Math.Round(Math.Min(1.0, Math.Max(0.0, totalScore)), 3);
        }

        private static ShapeComparisonResult CompareMatchedShapes(
            ShapeSnapshot src,
            ShapeSnapshot dst,
            int sourceSlideNumber,
            int destSlideNumber,
            double score,
            CompareSettings settings)
        {
            var diffTokens = TextDiffEngine.ComputeDiff(src.RawText, dst.RawText, settings.TextMatchingMode);
            bool hasTextDifference = !string.Equals(src.RawText?.Trim(), dst.RawText?.Trim(), StringComparison.Ordinal)
                                     || diffTokens.Any(t => t.DiffType != TextDiffType.Unchanged);
            string inlineDiff = TextDiffEngine.FormatInlineDiff(diffTokens);

            var formattingDiffs = FormattingComparisonEngine.CompareFormatting(src, dst, settings);

            // Check if moved
            float dx = Math.Abs(src.Left - dst.Left);
            float dy = Math.Abs(src.Top - dst.Top);
            bool isMoved = settings.ComparePosition && (dx > settings.PositionTolerancePoints || dy > settings.PositionTolerancePoints);

            // Determine match status
            ShapeMatchType matchType;
            if (hasTextDifference)
            {
                matchType = ShapeMatchType.Modified;
            }
            else if (isMoved)
            {
                matchType = ShapeMatchType.Moved;
            }
            else
            {
                matchType = ShapeMatchType.Matched;
            }

            string positionDiffText = string.Empty;
            if (isMoved)
            {
                positionDiffText = $"X: {src.Left:0}→{dst.Left:0} pt, Y: {src.Top:0}→{dst.Top:0} pt";
            }

            // Table comparison details
            string details = hasTextDifference ? "Text content modified" : (isMoved ? "Shape position moved" : (formattingDiffs.Count > 0 ? "Formatting modified" : "Content matched"));
            if (src.IsTable && dst.IsTable)
            {
                string tblDetails = CompareTables(src, dst);
                if (!string.IsNullOrEmpty(tblDetails))
                {
                    details = tblDetails;
                    matchType = ShapeMatchType.Modified;
                }
            }

            return new ShapeComparisonResult
            {
                SourceShapeId = src.Id,
                SourceShapeName = src.Name,
                DestinationShapeId = dst.Id,
                DestinationShapeName = dst.Name,
                MatchType = matchType,
                DifferenceCategory = hasTextDifference ? DifferenceCategory.Content : (isMoved ? DifferenceCategory.Layout : (formattingDiffs.Count > 0 ? DifferenceCategory.Formatting : DifferenceCategory.None)),
                SimilarityScore = score,
                SourceText = src.RawText,
                DestinationText = dst.RawText,
                TextDiff = inlineDiff,
                DiffTokens = diffTokens,
                FormattingDifferences = formattingDiffs,
                SourceSlideNumber = sourceSlideNumber,
                DestinationSlideNumber = destSlideNumber,
                DestinationShapeIndex = dst.Index,
                PositionDiff = positionDiffText,
                Details = details
            };
        }

        private static string CompareTables(ShapeSnapshot src, ShapeSnapshot dst)
        {
            var cellDiffs = new List<string>();
            int minRows = Math.Min(src.TableRowCount, dst.TableRowCount);
            int minCols = Math.Min(src.TableColumnCount, dst.TableColumnCount);

            for (int r = 0; r < minRows; r++)
            {
                for (int c = 0; c < minCols; c++)
                {
                    var srcCell = src.TableCells.FirstOrDefault(tc => tc.RowIndex == r && tc.ColumnIndex == c);
                    var dstCell = dst.TableCells.FirstOrDefault(tc => tc.RowIndex == r && tc.ColumnIndex == c);

                    string sText = srcCell?.Text ?? string.Empty;
                    string dText = dstCell?.Text ?? string.Empty;

                    if (!string.Equals(sText.Trim(), dText.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        cellDiffs.Add($"R{r + 1}C{c + 1}: '{sText.Trim()}' → '{dText.Trim()}'");
                    }
                }
            }

            if (src.TableRowCount != dst.TableRowCount || src.TableColumnCount != dst.TableColumnCount)
            {
                cellDiffs.Insert(0, $"Dimensions: {src.TableRowCount}x{src.TableColumnCount} → {dst.TableRowCount}x{dst.TableColumnCount}");
            }

            return string.Join("; ", cellDiffs.Take(5));
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OneAssist.SlideCompare.Models;
using OneAssist.SlideCompare.Services;

namespace OneAssist.SlideCompare.Matching
{
    /// <summary>
    /// Intelligently maps slides between source and destination presentations,
    /// prioritizing best 1-to-1 slide-to-slide matching, followed by secondary split/merged detection.
    /// </summary>
    public static class SlideMappingEngine
    {
        public static List<SlideComparisonResult> MapSlides(
            PresentationSnapshot sourcePres,
            PresentationSnapshot destPres,
            CompareSettings settings)
        {
            var results = new List<SlideComparisonResult>();
            var sourceSlides = sourcePres?.Slides ?? new List<SlideSnapshot>();
            var destSlides = destPres?.Slides ?? new List<SlideSnapshot>();
            settings = settings ?? new CompareSettings();

            if (!settings.CompareHiddenSlides)
            {
                sourceSlides = sourceSlides.Where(s => !s.IsHidden).ToList();
                destSlides = destSlides.Where(s => !s.IsHidden).ToList();
            }

            int sCount = sourceSlides.Count;
            int dCount = destSlides.Count;

            // 1. Calculate similarity matrix between all source and destination slides
            double[,] simMatrix = new double[sCount, dCount];
            for (int i = 0; i < sCount; i++)
            {
                for (int j = 0; j < dCount; j++)
                {
                    simMatrix[i, j] = CalculateSlideSimilarity(sourceSlides[i], destSlides[j], settings);
                }
            }

            var claimedSourceIndices = new HashSet<int>();
            var claimedDestIndices = new HashSet<int>();

            // 2. PHASE 1: HIGH-CONFIDENCE 1-TO-1 BEST MATCHING
            // Greedily pair up slides with high content correspondence (e.g. 1-1, 2-2, 3-4, 4-3).
            var highConfidencePairs = new List<Tuple<int, int, double>>();
            for (int i = 0; i < sCount; i++)
            {
                for (int j = 0; j < dCount; j++)
                {
                    double sim = simMatrix[i, j];
                    if (sim >= 0.72)
                    {
                        highConfidencePairs.Add(Tuple.Create(i, j, sim));
                    }
                }
            }

            highConfidencePairs.Sort((a, b) => b.Item3.CompareTo(a.Item3));

            foreach (var pair in highConfidencePairs)
            {
                int si = pair.Item1;
                int dj = pair.Item2;
                double score = pair.Item3;

                if (claimedSourceIndices.Contains(si) || claimedDestIndices.Contains(dj))
                    continue;

                claimedSourceIndices.Add(si);
                claimedDestIndices.Add(dj);

                results.Add(BuildSlideComparisonResult(sourceSlides[si], destSlides[dj], score, settings));
            }

            // 3. PHASE 2: SECONDARY SPLIT & MERGED DETECTION ON UNCLAIMED SLIDES
            // Detect 1-to-many Split slides
            for (int i = 0; i < sCount; i++)
            {
                if (claimedSourceIndices.Contains(i)) continue;

                var matchingDestList = new List<int>();
                for (int j = 0; j < dCount; j++)
                {
                    if (claimedDestIndices.Contains(j)) continue;

                    if (simMatrix[i, j] >= 0.20)
                    {
                        double overlap = CalculateContentContainment(destSlides[j], sourceSlides[i]);
                        if (overlap >= 0.55)
                        {
                            matchingDestList.Add(j);
                        }
                    }
                }

                if (matchingDestList.Count >= 2)
                {
                    claimedSourceIndices.Add(i);
                    foreach (var dj in matchingDestList)
                    {
                        claimedDestIndices.Add(dj);
                    }

                    var splitRes = new SlideComparisonResult
                    {
                        SourceSlideNumber = sourceSlides[i].SlideNumber,
                        SourceSlideNumbers = new List<int> { sourceSlides[i].SlideNumber },
                        DestinationSlideNumbers = matchingDestList.Select(dj => destSlides[dj].SlideNumber).ToList(),
                        MatchType = SlideMatchType.Split,
                        SimilarityScore = matchingDestList.Average(dj => simMatrix[i, dj]),
                        SourceTitle = sourceSlides[i].Title,
                        DestinationTitle = string.Join(" / ", matchingDestList.Select(dj => destSlides[dj].Title).Where(t => !string.IsNullOrEmpty(t))),
                        SummaryText = $"Source Slide {sourceSlides[i].SlideNumber} was split across Destination Slides {string.Join(", ", matchingDestList.Select(dj => destSlides[dj].SlideNumber))}."
                    };

                    var combinedDestShapes = matchingDestList.SelectMany(dj => destSlides[dj].GetAllShapesFlattened()).ToList();
                    splitRes.ShapeResults = ShapeMatchingEngine.MatchShapes(
                        sourceSlides[i].GetAllShapesFlattened(),
                        combinedDestShapes,
                        sourceSlides[i].SlideNumber,
                        splitRes.PrimaryDestinationSlide,
                        settings);

                    results.Add(splitRes);
                }
            }

            // Detect many-to-1 Merged slides
            for (int j = 0; j < dCount; j++)
            {
                if (claimedDestIndices.Contains(j)) continue;

                var matchingSourceList = new List<int>();
                for (int i = 0; i < sCount; i++)
                {
                    if (claimedSourceIndices.Contains(i)) continue;

                    if (simMatrix[i, j] >= 0.20)
                    {
                        double overlap = CalculateContentContainment(sourceSlides[i], destSlides[j]);
                        if (overlap >= 0.55)
                        {
                            matchingSourceList.Add(i);
                        }
                    }
                }

                if (matchingSourceList.Count >= 2)
                {
                    claimedDestIndices.Add(j);
                    foreach (var si in matchingSourceList)
                    {
                        claimedSourceIndices.Add(si);
                    }

                    var mergedRes = new SlideComparisonResult
                    {
                        SourceSlideNumber = sourceSlides[matchingSourceList[0]].SlideNumber,
                        SourceSlideNumbers = matchingSourceList.Select(si => sourceSlides[si].SlideNumber).ToList(),
                        DestinationSlideNumbers = new List<int> { destSlides[j].SlideNumber },
                        MatchType = SlideMatchType.Merged,
                        SimilarityScore = matchingSourceList.Average(si => simMatrix[si, j]),
                        SourceTitle = string.Join(" / ", matchingSourceList.Select(si => sourceSlides[si].Title).Where(t => !string.IsNullOrEmpty(t))),
                        DestinationTitle = destSlides[j].Title,
                        SummaryText = $"Source Slides {string.Join(", ", matchingSourceList.Select(si => sourceSlides[si].SlideNumber))} were merged into Destination Slide {destSlides[j].SlideNumber}."
                    };

                    var combinedSourceShapes = matchingSourceList.SelectMany(si => sourceSlides[si].GetAllShapesFlattened()).ToList();
                    mergedRes.ShapeResults = ShapeMatchingEngine.MatchShapes(
                        combinedSourceShapes,
                        destSlides[j].GetAllShapesFlattened(),
                        mergedRes.PrimarySourceSlide,
                        destSlides[j].SlideNumber,
                        settings);

                    results.Add(mergedRes);
                }
            }

            // 4. PHASE 3: MODERATE 1-TO-1 MATCHING ON REMAINING UNCLAIMED SLIDES
            var remainingPairs = new List<Tuple<int, int, double>>();
            for (int i = 0; i < sCount; i++)
            {
                if (claimedSourceIndices.Contains(i)) continue;
                for (int j = 0; j < dCount; j++)
                {
                    if (claimedDestIndices.Contains(j)) continue;
                    double sim = simMatrix[i, j];
                    if (sim >= 0.20)
                    {
                        remainingPairs.Add(Tuple.Create(i, j, sim));
                    }
                }
            }

            remainingPairs.Sort((a, b) => b.Item3.CompareTo(a.Item3));

            foreach (var pair in remainingPairs)
            {
                int si = pair.Item1;
                int dj = pair.Item2;
                double score = pair.Item3;

                if (claimedSourceIndices.Contains(si) || claimedDestIndices.Contains(dj))
                    continue;

                claimedSourceIndices.Add(si);
                claimedDestIndices.Add(dj);

                results.Add(BuildSlideComparisonResult(sourceSlides[si], destSlides[dj], score, settings));
            }

            // 5. UNMATCHED SOURCE SLIDES -> DELETED
            for (int i = 0; i < sCount; i++)
            {
                if (!claimedSourceIndices.Contains(i))
                {
                    var src = sourceSlides[i];
                    var deletedRes = new SlideComparisonResult
                    {
                        SourceSlideNumber = src.SlideNumber,
                        SourceSlideNumbers = new List<int> { src.SlideNumber },
                        DestinationSlideNumbers = new List<int>(),
                        MatchType = SlideMatchType.Deleted,
                        SimilarityScore = 0.0,
                        SourceTitle = src.Title,
                        DestinationTitle = string.Empty,
                        SummaryText = $"Slide {src.SlideNumber} was deleted in destination.",
                        ShapeResults = src.GetAllShapesFlattened().Select(s => new ShapeComparisonResult
                        {
                            SourceShapeId = s.Id,
                            SourceShapeName = s.Name,
                            MatchType = ShapeMatchType.Deleted,
                            SourceText = s.RawText,
                            TextDiff = !string.IsNullOrWhiteSpace(s.RawText) ? $"- {s.RawText.Trim()}" : string.Empty,
                            DiffTokens = !string.IsNullOrWhiteSpace(s.RawText)
                                ? new List<TextDiffToken> { new TextDiffToken { DiffType = TextDiffType.Deleted, Text = s.RawText } }
                                : new List<TextDiffToken>(),
                            SourceSlideNumber = src.SlideNumber,
                            Details = $"Deleted shape '{s.Name}'"
                        }).ToList()
                    };
                    results.Add(deletedRes);
                }
            }

            // 6. UNMATCHED DESTINATION SLIDES -> ADDED
            for (int j = 0; j < dCount; j++)
            {
                if (!claimedDestIndices.Contains(j))
                {
                    var dst = destSlides[j];
                    var addedRes = new SlideComparisonResult
                    {
                        SourceSlideNumber = 0,
                        SourceSlideNumbers = new List<int>(),
                        DestinationSlideNumbers = new List<int> { dst.SlideNumber },
                        MatchType = SlideMatchType.Added,
                        SimilarityScore = 0.0,
                        SourceTitle = string.Empty,
                        DestinationTitle = dst.Title,
                        SummaryText = $"Slide {dst.SlideNumber} was added in destination.",
                        ShapeResults = dst.GetAllShapesFlattened().Select(s => new ShapeComparisonResult
                        {
                            DestinationShapeId = s.Id,
                            DestinationShapeName = s.Name,
                            DestinationShapeIndex = s.Index,
                            MatchType = ShapeMatchType.Added,
                            DestinationText = s.RawText,
                            TextDiff = !string.IsNullOrWhiteSpace(s.RawText) ? $"+ {s.RawText.Trim()}" : string.Empty,
                            DiffTokens = !string.IsNullOrWhiteSpace(s.RawText)
                                ? new List<TextDiffToken> { new TextDiffToken { DiffType = TextDiffType.Added, Text = s.RawText } }
                                : new List<TextDiffToken>(),
                            DestinationSlideNumber = dst.SlideNumber,
                            Details = $"Added shape '{s.Name}'"
                        }).ToList()
                    };
                    results.Add(addedRes);
                }
            }

            // Sort results by primary Source slide number so pairings appear in sequential source order (1, 2, 3, 4)
            // Added destination slides appear at the end.
            results.Sort((a, b) =>
            {
                int orderA = a.PrimarySourceSlide > 0 ? a.PrimarySourceSlide : 10000 + a.PrimaryDestinationSlide;
                int orderB = b.PrimarySourceSlide > 0 ? b.PrimarySourceSlide : 10000 + b.PrimaryDestinationSlide;
                return orderA.CompareTo(orderB);
            });

            return results;
        }

        private static SlideComparisonResult BuildSlideComparisonResult(
            SlideSnapshot src,
            SlideSnapshot dst,
            double score,
            CompareSettings settings)
        {
            var shapeResults = ShapeMatchingEngine.MatchShapes(
                src.GetAllShapesFlattened(),
                dst.GetAllShapesFlattened(),
                src.SlideNumber,
                dst.SlideNumber,
                settings);

            bool hasShapeIssues = shapeResults.Any(s => s.IsIssue);
            bool isReordered = (src.SlideNumber != dst.SlideNumber);

            SlideMatchType matchType;
            if (isReordered)
            {
                matchType = SlideMatchType.Reordered;
            }
            else if (hasShapeIssues)
            {
                matchType = SlideMatchType.Modified;
            }
            else
            {
                matchType = SlideMatchType.Matched;
            }

            string summary;
            if (isReordered)
            {
                summary = hasShapeIssues
                    ? $"Slide moved from position {src.SlideNumber} to {dst.SlideNumber}, and content was modified."
                    : $"Slide moved from position {src.SlideNumber} to {dst.SlideNumber}.";
            }
            else
            {
                summary = hasShapeIssues ? "Slide content or formatting modified." : "Slide fully matched.";
            }

            var slideRes = new SlideComparisonResult
            {
                SourceSlideNumber = src.SlideNumber,
                SourceSlideNumbers = new List<int> { src.SlideNumber },
                DestinationSlideNumbers = new List<int> { dst.SlideNumber },
                MatchType = matchType,
                SimilarityScore = score,
                SourceTitle = src.Title,
                DestinationTitle = dst.Title,
                ShapeResults = shapeResults,
                SummaryText = summary
            };

            if (settings.CompareSpeakerNotes && (!string.IsNullOrEmpty(src.SpeakerNotes) || !string.IsNullOrEmpty(dst.SpeakerNotes)))
            {
                slideRes.SpeakerNotesDiff = TextDiffEngine.FormatInlineDiff(
                    TextDiffEngine.ComputeDiff(src.SpeakerNotes, dst.SpeakerNotes, settings.TextMatchingMode));
            }

            return slideRes;
        }

        /// <summary>
        /// Calculates comprehensive similarity between two slides based on title, body, tables, and shapes.
        /// </summary>
        public static double CalculateSlideSimilarity(SlideSnapshot src, SlideSnapshot dst, CompareSettings settings)
        {
            if (src == null || dst == null) return 0.0;

            // 1. Exact Fingerprint Match
            if (!string.IsNullOrEmpty(src.Fingerprint) && string.Equals(src.Fingerprint, dst.Fingerprint, StringComparison.OrdinalIgnoreCase))
            {
                return 1.0;
            }

            // 2. Title Similarity
            bool srcHasTitle = !string.IsNullOrWhiteSpace(src.Title);
            bool dstHasTitle = !string.IsNullOrWhiteSpace(dst.Title);
            double titleSim = 0.0;
            double titleWeight = 0.25;

            if (srcHasTitle && dstHasTitle)
            {
                titleSim = SimilarityCalculator.CalculateLevenshteinSimilarity(src.NormalizedTitle, dst.NormalizedTitle);
            }
            else if (!srcHasTitle && !dstHasTitle)
            {
                titleWeight = 0.0;
            }
            else
            {
                titleSim = 0.0;
            }

            // 3. Complete Text Similarity across all shapes on the slide
            string srcAllText = SimilarityCalculator.NormalizeText(GetAllSlideText(src));
            string dstAllText = SimilarityCalculator.NormalizeText(GetAllSlideText(dst));

            bool srcHasText = !string.IsNullOrWhiteSpace(srcAllText);
            bool dstHasText = !string.IsNullOrWhiteSpace(dstAllText);
            double textSim = 0.0;
            double textWeight = titleWeight > 0.0 ? 0.45 : 0.65;

            if (srcHasText && dstHasText)
            {
                double dice = SimilarityCalculator.CalculateTextDiceSimilarity(srcAllText, dstAllText);
                double jaccard = CalculateTokenJaccard(srcAllText, dstAllText);
                textSim = Math.Max(dice, jaccard);
            }
            else if (!srcHasText && !dstHasText)
            {
                textSim = 1.0;
            }
            else
            {
                textSim = 0.0;
            }

            // 4. Shape Type & Count Signature
            var srcShapes = src.GetAllShapesFlattened();
            var dstShapes = dst.GetAllShapesFlattened();
            double shapeSim = 0.0;
            double countSim = 0.0;
            double shapeWeight = 0.0;
            double countWeight = 0.0;
            int maxShapes = Math.Max(srcShapes.Count, dstShapes.Count);

            if (maxShapes > 0)
            {
                shapeWeight = 0.20;
                countWeight = 0.10;
                countSim = 1.0 - (Math.Abs(srcShapes.Count - dstShapes.Count) / (double)maxShapes);

                if (srcShapes.Count > 0 && dstShapes.Count > 0)
                {
                    double totalShapeScore = 0.0;
                    foreach (var s in srcShapes)
                    {
                        double bestS = 0.0;
                        foreach (var d in dstShapes)
                        {
                            double sSim = ShapeMatchingEngine.CalculateShapeSimilarity(s, d, settings);
                            if (sSim > bestS) bestS = sSim;
                        }
                        totalShapeScore += bestS;
                    }
                    shapeSim = totalShapeScore / srcShapes.Count;
                }
            }

            double totalWeight = titleWeight + textWeight + shapeWeight + countWeight;
            if (totalWeight <= 0.0) return 0.0;

            double score = ((titleSim * titleWeight) + (textSim * textWeight) + (shapeSim * shapeWeight) + (countSim * countWeight)) / totalWeight;

            // 5. Positional alignment bonus: if same slide number and good content similarity, grant small tie-breaker bonus
            if (src.SlideNumber == dst.SlideNumber && score >= 0.25)
            {
                score += 0.05;
            }

            return Math.Round(Math.Min(1.0, Math.Max(0.0, score)), 3);
        }

        public static string GetAllSlideText(SlideSnapshot slide)
        {
            if (slide == null) return string.Empty;
            var sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(slide.Title)) sb.AppendLine(slide.Title);
            if (!string.IsNullOrWhiteSpace(slide.BodyText)) sb.AppendLine(slide.BodyText);

            var shapes = slide.GetAllShapesFlattened();
            foreach (var shp in shapes)
            {
                if (shp.HasText && !string.IsNullOrWhiteSpace(shp.RawText))
                {
                    string t = shp.RawText.Trim();
                    if (!t.Equals(slide.Title?.Trim(), StringComparison.OrdinalIgnoreCase) &&
                        !t.Equals(slide.BodyText?.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        sb.AppendLine(t);
                    }
                }
            }
            return sb.ToString();
        }

        private static double CalculateTokenJaccard(string textA, string textB)
        {
            if (string.IsNullOrEmpty(textA) || string.IsNullOrEmpty(textB)) return 0.0;
            var setA = new HashSet<string>(textA.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
            var setB = new HashSet<string>(textB.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
            if (setA.Count == 0 || setB.Count == 0) return 0.0;
            int intersection = setA.Count(w => setB.Contains(w));
            int union = setA.Count + setB.Count - intersection;
            return union > 0 ? (double)intersection / union : 0.0;
        }

        /// <summary>
        /// Measures what fraction of words from subSlide are contained in containerSlide.
        /// Useful for split and merged slide detection.
        /// </summary>
        private static double CalculateContentContainment(SlideSnapshot subSlide, SlideSnapshot containerSlide)
        {
            string subText = SimilarityCalculator.NormalizeText(GetAllSlideText(subSlide));
            string contText = SimilarityCalculator.NormalizeText(GetAllSlideText(containerSlide));

            if (string.IsNullOrEmpty(subText) || string.IsNullOrEmpty(contText)) return 0.0;

            var subWords = subText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (subWords.Length == 0) return 0.0;

            var contWordSet = new HashSet<string>(contText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
            int found = subWords.Count(w => contWordSet.Contains(w));

            return (double)found / subWords.Length;
        }
    }
}

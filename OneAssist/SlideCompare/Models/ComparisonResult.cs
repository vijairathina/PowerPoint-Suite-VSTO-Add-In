using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace OneAssist.SlideCompare.Models
{
    /// <summary>
    /// Represents a discrete text diff segment (added, deleted, modified, or unchanged).
    /// </summary>
    [DataContract]
    public class TextDiffToken
    {
        [DataMember]
        public TextDiffType DiffType { get; set; }

        [DataMember]
        public string Text { get; set; } = string.Empty;

        public override string ToString()
        {
            switch (DiffType)
            {
                case TextDiffType.Added: return $"+[{Text}]";
                case TextDiffType.Deleted: return $"-[{Text}]";
                case TextDiffType.Modified: return $"~[{Text}]";
                case TextDiffType.Moved: return $"⇄[{Text}]";
                default: return Text;
            }
        }
    }

    /// <summary>
    /// Details for formatting changes between source and destination shapes.
    /// </summary>
    [DataContract]
    public class FormattingDifferenceItem
    {
        [DataMember]
        public string PropertyName { get; set; } = string.Empty;

        [DataMember]
        public string SourceValue { get; set; } = string.Empty;

        [DataMember]
        public string DestinationValue { get; set; } = string.Empty;

        public override string ToString()
        {
            return $"{PropertyName}: '{SourceValue}' → '{DestinationValue}'";
        }
    }

    /// <summary>
    /// Comparison result for an individual shape.
    /// </summary>
    [DataContract]
    public class ShapeComparisonResult
    {
        [DataMember]
        public int SourceShapeId { get; set; }

        [DataMember]
        public string SourceShapeName { get; set; } = string.Empty;

        [DataMember]
        public int DestinationShapeId { get; set; }

        [DataMember]
        public string DestinationShapeName { get; set; } = string.Empty;

        [DataMember]
        public ShapeMatchType MatchType { get; set; }

        [DataMember]
        public DifferenceCategory DifferenceCategory { get; set; }

        [DataMember]
        public double SimilarityScore { get; set; }

        [DataMember]
        public string SourceText { get; set; } = string.Empty;

        [DataMember]
        public string DestinationText { get; set; } = string.Empty;

        [DataMember]
        public string TextDiff { get; set; } = string.Empty;

        [DataMember]
        public List<TextDiffToken> DiffTokens { get; set; } = new List<TextDiffToken>();

        [DataMember]
        public List<FormattingDifferenceItem> FormattingDifferences { get; set; } = new List<FormattingDifferenceItem>();

        [DataMember]
        public int SourceSlideNumber { get; set; }

        [DataMember]
        public int DestinationSlideNumber { get; set; }

        [DataMember]
        public int DestinationShapeIndex { get; set; }

        [DataMember]
        public string Details { get; set; } = string.Empty;

        // Position changes if moved
        [DataMember]
        public string PositionDiff { get; set; } = string.Empty;

        public string DisplayName => !string.IsNullOrEmpty(DestinationShapeName)
            ? DestinationShapeName
            : (!string.IsNullOrEmpty(SourceShapeName) ? SourceShapeName : $"Shape #{SourceShapeId}");

        public bool HasSourceText => !string.IsNullOrWhiteSpace(SourceText);
        public bool HasDestinationText => !string.IsNullOrWhiteSpace(DestinationText);

        public bool HasTextChanges =>
            (DiffTokens != null && DiffTokens.Any(t => t.DiffType != TextDiffType.Unchanged))
            || (!string.IsNullOrEmpty(SourceText) && !string.IsNullOrEmpty(DestinationText) && SourceText.Trim() != DestinationText.Trim())
            || (string.IsNullOrEmpty(SourceText) != string.IsNullOrEmpty(DestinationText));

        public bool HasFormattingChanges => FormattingDifferences != null && FormattingDifferences.Count > 0;
        public bool IsIssue => MatchType != ShapeMatchType.Matched || HasTextChanges || HasFormattingChanges;

        public bool IsDeleted => MatchType == ShapeMatchType.Deleted;
        public bool IsAdded => MatchType == ShapeMatchType.Added;
        public bool IsModified => MatchType == ShapeMatchType.Modified || HasTextChanges;
        public bool IsMoved => MatchType == ShapeMatchType.Moved || !string.IsNullOrEmpty(PositionDiff);
        public bool HasPositionDiff => !string.IsNullOrEmpty(PositionDiff);
        public bool IsMatched => MatchType == ShapeMatchType.Matched || (!HasTextChanges && !IsDeleted && !IsAdded && !IsMoved);
        public bool IsContentMatched => MatchType == ShapeMatchType.Matched || (!HasTextChanges && !IsDeleted && !IsAdded);

        public string ChangesSummary => Matching.TextDiffEngine.FormatChangesOnly(DiffTokens);
        public bool HasChangesSummary => !string.IsNullOrWhiteSpace(ChangesSummary);

        public string StatusBadgeText
        {
            get
            {
                switch (MatchType)
                {
                    case ShapeMatchType.Matched: return HasFormattingChanges ? "FORMAT MODIFIED" : "MATCHED";
                    case ShapeMatchType.Modified: return "MODIFIED";
                    case ShapeMatchType.Added: return "ADDED";
                    case ShapeMatchType.Deleted: return "DELETED";
                    case ShapeMatchType.Moved: return "MOVED";
                    default: return MatchType.ToString().ToUpperInvariant();
                }
            }
        }

        public string StatusIcon
        {
            get
            {
                switch (MatchType)
                {
                    case ShapeMatchType.Matched: return HasFormattingChanges ? "✎" : "✓";
                    case ShapeMatchType.Modified: return "✎";
                    case ShapeMatchType.Added: return "+";
                    case ShapeMatchType.Deleted: return "−";
                    case ShapeMatchType.Moved: return "↔";
                    default: return "•";
                }
            }
        }
    }

    /// <summary>
    /// Comparison result for a slide (can be 1-to-1, 1-to-many Split, many-to-1 Merged, etc.).
    /// </summary>
    [DataContract]
    public class SlideComparisonResult
    {
        [DataMember]
        public int SourceSlideNumber { get; set; }

        [DataMember]
        public List<int> DestinationSlideNumbers { get; set; } = new List<int>();

        [DataMember]
        public List<int> SourceSlideNumbers { get; set; } = new List<int>();

        [DataMember]
        public SlideMatchType MatchType { get; set; }

        [DataMember]
        public double SimilarityScore { get; set; }

        [DataMember]
        public string SourceTitle { get; set; } = string.Empty;

        [DataMember]
        public string DestinationTitle { get; set; } = string.Empty;

        [DataMember]
        public List<ShapeComparisonResult> ShapeResults { get; set; } = new List<ShapeComparisonResult>();

        [DataMember]
        public string SpeakerNotesDiff { get; set; } = string.Empty;

        [DataMember]
        public string SummaryText { get; set; } = string.Empty;

        public int PrimaryDestinationSlide => DestinationSlideNumbers != null && DestinationSlideNumbers.Count > 0 ? DestinationSlideNumbers[0] : 0;
        public int PrimarySourceSlide => SourceSlideNumbers != null && SourceSlideNumbers.Count > 0 ? SourceSlideNumbers[0] : SourceSlideNumber;

        public bool HasIssues => MatchType != SlideMatchType.Matched || (ShapeResults != null && ShapeResults.Any(s => s.IsIssue));

        public bool ShowOnlyIssues { get; set; } = false;

        public bool IsMatched => MatchType == SlideMatchType.Matched && !HasIssues;

        public IEnumerable<ShapeComparisonResult> DisplayShapes
        {
            get
            {
                if (ShapeResults == null) return Enumerable.Empty<ShapeComparisonResult>();
                return ShowOnlyIssues ? ShapeResults.Where(s => s.IsIssue) : ShapeResults;
            }
        }

        public string TitleDisplay
        {
            get
            {
                string titlePart = string.Empty;
                if (!string.IsNullOrEmpty(DestinationTitle) && !string.IsNullOrEmpty(SourceTitle))
                {
                    titlePart = string.Equals(DestinationTitle, SourceTitle, StringComparison.OrdinalIgnoreCase)
                        ? $" • {DestinationTitle}"
                        : $" • {SourceTitle} → {DestinationTitle}";
                }
                else if (!string.IsNullOrEmpty(DestinationTitle))
                {
                    titlePart = $" • {DestinationTitle}";
                }
                else if (!string.IsNullOrEmpty(SourceTitle))
                {
                    titlePart = $" • {SourceTitle}";
                }

                switch (MatchType)
                {
                    case SlideMatchType.Split:
                        return $"Slide {SourceSlideNumber} → Slides {string.Join(", ", DestinationSlideNumbers)}{titlePart}";
                    case SlideMatchType.Merged:
                        return $"Slides {string.Join(", ", SourceSlideNumbers)} → Slide {PrimaryDestinationSlide}{titlePart}";
                    case SlideMatchType.Added:
                        return $"Slide {PrimaryDestinationSlide}{titlePart}";
                    case SlideMatchType.Deleted:
                        return $"Slide {SourceSlideNumber}{titlePart}";
                    case SlideMatchType.Reordered:
                        return $"Slide {SourceSlideNumber} → Slide {PrimaryDestinationSlide}{titlePart}";
                    case SlideMatchType.Moved:
                        return $"Slide {SourceSlideNumber} → Slide {PrimaryDestinationSlide}{titlePart}";
                    default:
                        return $"Slide {SourceSlideNumber} → Slide {PrimaryDestinationSlide}{titlePart}";
                }
            }
        }

        public string StatusBadgeText => MatchType.ToString().ToUpperInvariant();

        public string StatusIcon
        {
            get
            {
                switch (MatchType)
                {
                    case SlideMatchType.Matched: return "✓";
                    case SlideMatchType.Modified: return "✎";
                    case SlideMatchType.Added: return "+";
                    case SlideMatchType.Deleted: return "−";
                    case SlideMatchType.Moved: return "↔";
                    case SlideMatchType.Reordered: return "⇄";
                    case SlideMatchType.Split: return "⇢";
                    case SlideMatchType.Merged: return "⇠";
                    default: return "•";
                }
            }
        }

        private string _cachedSearchableText;

        public bool MatchesSearch(string queryLower)
        {
            if (string.IsNullOrWhiteSpace(queryLower)) return true;

            if (_cachedSearchableText == null)
            {
                var sb = new System.Text.StringBuilder();
                sb.Append(SourceSlideNumber).Append(' ');
                sb.Append(PrimaryDestinationSlide).Append(' ');
                if (DestinationSlideNumbers != null) sb.Append(string.Join(" ", DestinationSlideNumbers)).Append(' ');
                if (SourceSlideNumbers != null) sb.Append(string.Join(" ", SourceSlideNumbers)).Append(' ');
                if (!string.IsNullOrEmpty(SourceTitle)) sb.Append(SourceTitle).Append(' ');
                if (!string.IsNullOrEmpty(DestinationTitle)) sb.Append(DestinationTitle).Append(' ');
                if (!string.IsNullOrEmpty(SummaryText)) sb.Append(SummaryText).Append(' ');
                if (!string.IsNullOrEmpty(SpeakerNotesDiff)) sb.Append(SpeakerNotesDiff).Append(' ');

                if (ShapeResults != null)
                {
                    foreach (var sh in ShapeResults)
                    {
                        if (!string.IsNullOrEmpty(sh.SourceShapeName)) sb.Append(sh.SourceShapeName).Append(' ');
                        if (!string.IsNullOrEmpty(sh.DestinationShapeName)) sb.Append(sh.DestinationShapeName).Append(' ');
                        if (!string.IsNullOrEmpty(sh.SourceText)) sb.Append(sh.SourceText).Append(' ');
                        if (!string.IsNullOrEmpty(sh.DestinationText)) sb.Append(sh.DestinationText).Append(' ');
                        if (!string.IsNullOrEmpty(sh.TextDiff)) sb.Append(sh.TextDiff).Append(' ');
                        if (!string.IsNullOrEmpty(sh.Details)) sb.Append(sh.Details).Append(' ');
                        if (!string.IsNullOrEmpty(sh.PositionDiff)) sb.Append(sh.PositionDiff).Append(' ');
                    }
                }

                _cachedSearchableText = sb.ToString().ToLowerInvariant();
            }

            return _cachedSearchableText.Contains(queryLower);
        }
    }

    /// <summary>
    /// Master comparison result for the entire presentation.
    /// </summary>
    [DataContract]
    public class PresentationComparisonResult
    {
        [DataMember]
        public string SourcePresentationName { get; set; } = string.Empty;

        [DataMember]
        public string DestinationPresentationName { get; set; } = string.Empty;

        [DataMember]
        public DateTime ComparedAt { get; set; } = DateTime.UtcNow;

        [DataMember]
        public int TotalSourceSlides { get; set; }

        [DataMember]
        public int TotalDestinationSlides { get; set; }

        [DataMember]
        public int MatchedCount { get; set; }

        [DataMember]
        public int MatchedContentCount { get; set; }

        [DataMember]
        public int TotalContentCount { get; set; }

        [DataMember]
        public int MatchedSlidesCount { get; set; }

        [DataMember]
        public int ModifiedCount { get; set; }

        [DataMember]
        public int AddedCount { get; set; }

        [DataMember]
        public int DeletedCount { get; set; }

        [DataMember]
        public int MovedCount { get; set; }

        [DataMember]
        public int ReorderedCount { get; set; }

        [DataMember]
        public int SplitCount { get; set; }

        [DataMember]
        public int MergedCount { get; set; }

        [DataMember]
        public List<SlideComparisonResult> SlideResults { get; set; } = new List<SlideComparisonResult>();

        public int TotalIssuesCount => ModifiedCount + AddedCount + DeletedCount + MovedCount + ReorderedCount + SplitCount + MergedCount;
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Alignment
{
    public class InconsistentSpacingRule : IQualityRule
    {
        public string RuleId => "ALIGN_INCONSISTENT_SPACING";
        public string Name => "Inconsistent Spacing";
        public string Description => "Detects sequences of 3 or more aligned cards, boxes or objects with uneven gaps between them.";
        public RuleCategory Category => RuleCategory.Alignment;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => settings?.EnableSpacingCheck ?? true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null || context.SlideShapes.Count < 3) return issues;

            float tol = context.Settings?.SpacingInconsistencyTolerancePoints ?? 2.0f;
            var shapes = context.SlideShapes;

            // Group shapes that share roughly the same Top (horizontal row)
            var horizontalRows = shapes.GroupBy(s => (int)Math.Round(s.Top / 15.0f))
                                       .Where(g => g.Count() >= 3);

            foreach (var group in horizontalRows)
            {
                var rowShapes = group.OrderBy(s => s.Left).ToList();
                var gaps = new List<float>();
                for (int i = 0; i < rowShapes.Count - 1; i++)
                {
                    float gap = rowShapes[i + 1].Left - (rowShapes[i].Left + rowShapes[i].Width);
                    gaps.Add(gap);
                }

                float minGap = gaps.Min();
                float maxGap = gaps.Max();

                if (maxGap - minGap > tol && minGap > 0)
                {
                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = Name,
                        Category = Category,
                        Severity = DefaultSeverity,
                        SlideIndex = context.SlideIndex,
                        ShapeId = rowShapes[1].Id,
                        ShapeName = $"Row of {rowShapes.Count} objects",
                        Problem = $"Gaps between shapes in this row vary between {minGap:F1}pt and {maxGap:F1}pt.",
                        Expected = $"Uniform {gaps.Average():F1}pt spacing",
                        Actual = $"{minGap:F1}pt to {maxGap:F1}pt",
                        Recommendation = "Distribute shapes horizontally to equalize spacing across the row.",
                        CanAutoFix = true,
                        FixPayload = rowShapes.Select(s => s.Id).ToList()
                    });
                }
            }

            return issues;
        }

        public bool Fix(QualityIssue issue, PowerPoint.Application app, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (issue == null || app == null) return false;

            try
            {
                var pres = app.ActivePresentation;
                var slide = pres.Slides[issue.SlideIndex];
                var shapeIds = issue.FixPayload as List<int>;
                if (shapeIds == null || shapeIds.Count < 3)
                {
                    errorMessage = "Invalid shape group for distribution.";
                    return false;
                }

                var matchedShapes = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (shapeIds.Contains(s.Id))
                    {
                        matchedShapes.Add(s);
                    }
                }

                if (matchedShapes.Count >= 3)
                {
                    OneAssist.Assistant.Services.AlignmentService.DistributeShapesHorizontally(matchedShapes);
                    issue.IsFixed = true;
                    return true;
                }

                errorMessage = "Could not locate all shapes in row.";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }
    }
}

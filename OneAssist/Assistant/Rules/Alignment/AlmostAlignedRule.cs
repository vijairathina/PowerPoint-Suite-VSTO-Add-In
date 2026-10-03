using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Alignment
{
    public class AlmostAlignedRule : IQualityRule
    {
        public string RuleId => "ALIGN_ALMOST_ALIGNED";
        public string Name => "Almost-Aligned Objects";
        public string Description => "Detects objects that are slightly misaligned by 1 to 3 points, creating visual sloppiness.";
        public RuleCategory Category => RuleCategory.Alignment;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => settings?.EnableAlignmentCheck ?? true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null || context.SlideShapes.Count < 2) return issues;

            float tol = context.Settings?.AlmostAlignedTolerancePoints ?? 3.0f;
            var shapes = context.SlideShapes;

            for (int i = 0; i < shapes.Count - 1; i++)
            {
                if (context.CancellationToken.IsCancellationRequested) break;
                var s1 = shapes[i];

                for (int j = i + 1; j < shapes.Count; j++)
                {
                    var s2 = shapes[j];

                    // Check horizontal alignment (Top edges)
                    float dTop = Math.Abs(s1.Top - s2.Top);
                    if (dTop > 0.05f && dTop <= tol)
                    {
                        issues.Add(new QualityIssue
                        {
                            RuleId = RuleId,
                            RuleTitle = Name,
                            Category = Category,
                            Severity = DefaultSeverity,
                            SlideIndex = context.SlideIndex,
                            ShapeId = s2.Id,
                            ShapeName = $"{s2.Name} (with {s1.Name})",
                            Problem = $"Shapes are almost aligned on Top, differing by only {dTop:F1}pt.",
                            Expected = $"Top: {s1.Top:F1}pt",
                            Actual = $"Top: {s2.Top:F1}pt",
                            Recommendation = $"Snap top edge of '{s2.Name}' to match '{s1.Name}' at {s1.Top:F1}pt.",
                            CanAutoFix = true,
                            FixPayload = s1.Top,
                            BeforeStateSnapshot = s2.Top
                        });
                        break;
                    }

                    // Check vertical alignment (Left edges)
                    float dLeft = Math.Abs(s1.Left - s2.Left);
                    if (dLeft > 0.05f && dLeft <= tol)
                    {
                        issues.Add(new QualityIssue
                        {
                            RuleId = RuleId,
                            RuleTitle = Name,
                            Category = Category,
                            Severity = DefaultSeverity,
                            SlideIndex = context.SlideIndex,
                            ShapeId = s2.Id,
                            ShapeName = $"{s2.Name} (with {s1.Name})",
                            Problem = $"Shapes are almost aligned on Left, differing by only {dLeft:F1}pt.",
                            Expected = $"Left: {s1.Left:F1}pt",
                            Actual = $"Left: {s2.Left:F1}pt",
                            Recommendation = $"Snap left edge of '{s2.Name}' to match '{s1.Name}' at {s1.Left:F1}pt.",
                            CanAutoFix = true,
                            FixPayload = s1.Left,
                            BeforeStateSnapshot = s2.Left
                        });
                        break;
                    }
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
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (s.Id == issue.ShapeId)
                    {
                        if (issue.Expected.StartsWith("Top:") && issue.FixPayload is float targetTop)
                        {
                            s.Top = targetTop;
                            issue.IsFixed = true;
                            return true;
                        }
                        if (issue.Expected.StartsWith("Left:") && issue.FixPayload is float targetLeft)
                        {
                            s.Left = targetLeft;
                            issue.IsFixed = true;
                            return true;
                        }
                    }
                }
                errorMessage = "Target shape not found.";
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

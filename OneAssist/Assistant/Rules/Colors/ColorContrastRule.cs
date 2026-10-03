using System;
using System.Collections.Generic;
using System.Drawing;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Colors
{
    public class ColorContrastRule : IQualityRule
    {
        public string RuleId => "COLOR_CONTRAST";
        public string Name => "Low Text/Background Contrast";
        public string Description => "Checks text color against shape background to ensure compliance with WCAG AA readability (>= 4.5:1).";
        public RuleCategory Category => RuleCategory.Colors;
        public IssueSeverity DefaultSeverity => IssueSeverity.Error;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => settings?.EnableColorContrastCheck ?? true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null) return issues;

            double minRatio = context.Settings?.MinWcagContrastRatio ?? 4.5;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;
                if (!SafeInteropHelper.HasText(shape)) continue;

                Color? textColor = SafeInteropHelper.GetTextColor(shape);
                if (!textColor.HasValue) continue;

                Color? bgColor = SafeInteropHelper.GetFillColor(shape) ?? Color.White; // Default to white slide canvas

                double ratio = ColorEngine.CalculateContrastRatio(textColor.Value, bgColor.Value);
                if (ratio < minRatio)
                {
                    // Suggest high contrast text (either black or white depending on bg)
                    double contrastWithWhite = ColorEngine.CalculateContrastRatio(Color.White, bgColor.Value);
                    Color proposed = contrastWithWhite >= 4.5 ? Color.White : Color.Black;

                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = Name,
                        Category = Category,
                        Severity = DefaultSeverity,
                        SlideIndex = context.SlideIndex,
                        ShapeId = shape.Id,
                        ShapeName = shape.Name,
                        Problem = $"Text contrast ratio {ratio:F1}:1 is below WCAG AA standard ({minRatio}:1).",
                        Expected = $">= {minRatio}:1",
                        Actual = $"{ratio:F1}:1 (Text: {SafeInteropHelper.ColorToHex(textColor.Value)}, Bg: {SafeInteropHelper.ColorToHex(bgColor.Value)})",
                        Recommendation = $"Switch text color to {SafeInteropHelper.ColorToHex(proposed)} for crystal-clear readability.",
                        CanAutoFix = true,
                        FixPayload = proposed,
                        BeforeStateSnapshot = textColor.Value
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
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (s.Id == issue.ShapeId && SafeInteropHelper.HasText(s))
                    {
                        if (issue.FixPayload is Color newColor)
                        {
                            s.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(newColor);
                            issue.IsFixed = true;
                            return true;
                        }
                    }
                }
                errorMessage = "Target text shape not found.";
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

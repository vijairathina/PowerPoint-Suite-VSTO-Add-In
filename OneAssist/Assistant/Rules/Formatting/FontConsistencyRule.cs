using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Formatting
{
    public class FontConsistencyRule : IQualityRule
    {
        public string RuleId => "FMT_FONT_CONSISTENCY";
        public string Name => "Font Consistency";
        public string Description => "Flags text shapes using fonts that diverge from the approved brand or primary presentation font.";
        public RuleCategory Category => RuleCategory.Typography;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => settings?.EnableFontConsistencyCheck ?? true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null || context.Settings?.ActivePreset == null) return issues;

            string expectedPrimary = context.Settings.ActivePreset.PrimaryFont;
            string expectedHeading = context.Settings.ActivePreset.HeadingFont;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;
                if (!SafeInteropHelper.HasText(shape)) continue;

                string currentFont = SafeInteropHelper.GetFontName(shape);
                if (string.IsNullOrEmpty(currentFont)) continue;

                // Check against approved fonts
                bool matchesPrimary = string.Equals(currentFont, expectedPrimary, StringComparison.OrdinalIgnoreCase);
                bool matchesHeading = string.Equals(currentFont, expectedHeading, StringComparison.OrdinalIgnoreCase);

                if (!matchesPrimary && !matchesHeading)
                {
                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = Name,
                        Category = Category,
                        Severity = DefaultSeverity,
                        SlideIndex = context.SlideIndex,
                        ShapeId = shape.Id,
                        ShapeName = shape.Name,
                        Problem = $"Non-standard font '{currentFont}' used.",
                        Expected = expectedPrimary,
                        Actual = currentFont,
                        Recommendation = $"Switch font from '{currentFont}' to approved brand font '{expectedPrimary}'.",
                        CanAutoFix = true,
                        FixPayload = expectedPrimary,
                        BeforeStateSnapshot = currentFont
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
                    if (s.Id == issue.ShapeId)
                    {
                        string targetFont = issue.FixPayload as string ?? "Aptos";
                        TypographyEngine.ApplyFont(s, targetFont);
                        issue.IsFixed = true;
                        return true;
                    }
                }
                errorMessage = "Target shape not found on slide.";
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

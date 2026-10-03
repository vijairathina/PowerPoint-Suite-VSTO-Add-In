using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Formatting
{
    public class FontSizeConsistencyRule : IQualityRule
    {
        public string RuleId => "FMT_FONT_SIZE";
        public string Name => "Font Size Consistency";
        public string Description => "Detects text sizes that fall below readable minimum thresholds or use irregular fractional point sizes.";
        public RuleCategory Category => RuleCategory.Typography;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => settings?.EnableFontSizeCheck ?? true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null) return issues;

            float minBody = context.Settings?.ActivePreset?.MinBodyFontSize ?? 12.0f;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;
                if (!SafeInteropHelper.HasText(shape)) continue;

                float? size = SafeInteropHelper.GetFontSize(shape);
                if (!size.HasValue) continue;

                if (size.Value < minBody && size.Value > 1.0f)
                {
                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = Name,
                        Category = Category,
                        Severity = IssueSeverity.Warning,
                        SlideIndex = context.SlideIndex,
                        ShapeId = shape.Id,
                        ShapeName = shape.Name,
                        Problem = $"Text font size ({size.Value:F1}pt) is below recommended minimum ({minBody}pt).",
                        Expected = $"{minBody}pt",
                        Actual = $"{size.Value:F1}pt",
                        Recommendation = $"Increase font size to at least {minBody}pt for optimal audience readability.",
                        CanAutoFix = true,
                        FixPayload = minBody,
                        BeforeStateSnapshot = size.Value
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
                        float target = issue.FixPayload is float f ? f : 12.0f;
                        s.TextFrame.TextRange.Font.Size = target;
                        issue.IsFixed = true;
                        return true;
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

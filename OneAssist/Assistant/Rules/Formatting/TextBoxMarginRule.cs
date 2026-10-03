using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Formatting
{
    public class TextBoxMarginRule : IQualityRule
    {
        public string RuleId => "FMT_TEXT_MARGIN";
        public string Name => "Text Box Padding & Margins";
        public string Description => "Detects irregular or zeroed text box internal margins causing text to cling to borders.";
        public RuleCategory Category => RuleCategory.Formatting;
        public IssueSeverity DefaultSeverity => IssueSeverity.Info;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => settings?.EnableMarginCheck ?? true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null) return issues;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;
                if (!SafeInteropHelper.HasText(shape)) continue;

                try
                {
                    var tf = shape.TextFrame;
                    if (tf == null) continue;

                    // If shape has a visible fill or border, but zero margins, text touches the edge
                    bool hasFillOrLine = (shape.Fill != null && shape.Fill.Visible == Office.MsoTriState.msoTrue) ||
                                         (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue);

                    if (hasFillOrLine && (tf.MarginLeft < 2.0f || tf.MarginTop < 2.0f))
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
                            Problem = "Text box has zero or cramped internal padding inside a framed shape.",
                            Expected = "Standard padding (7.2pt)",
                            Actual = $"L:{tf.MarginLeft:F1}pt, T:{tf.MarginTop:F1}pt",
                            Recommendation = "Apply comfortable internal padding (7.2pt / 0.1in) so text doesn't clash with borders.",
                            CanAutoFix = true,
                            FixPayload = 7.2f
                        });
                    }
                }
                catch { }
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
                        float margin = issue.FixPayload is float f ? f : 7.2f;
                        TypographyEngine.NormalizeMargins(s, margin);
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

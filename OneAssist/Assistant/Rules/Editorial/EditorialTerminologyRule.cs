using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Editorial
{
    public class EditorialTerminologyRule : IQualityRule
    {
        public string RuleId => "EDIT_TERMINOLOGY";
        public string Name => "Corporate Terminology & Spelling";
        public string Description => "Checks presentation content against the corporate dictionary for non-standard terminology or spelling.";
        public RuleCategory Category => RuleCategory.Editorial;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null || context.Settings?.TerminologyReplacements == null) return issues;

            var dict = context.Settings.TerminologyReplacements;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;
                if (!SafeInteropHelper.HasText(shape)) continue;

                string text = SafeInteropHelper.GetText(shape);
                var results = EditorialEngine.CheckTerminology(text, dict);

                foreach (var res in results)
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
                        Problem = res.Problem,
                        Expected = res.Expected,
                        Actual = res.Actual,
                        Recommendation = res.Recommendation,
                        CanAutoFix = res.CanAutoFix,
                        FixPayload = res.FixedText,
                        BeforeStateSnapshot = text
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
                        if (issue.FixPayload is string fixedText)
                        {
                            s.TextFrame.TextRange.Text = fixedText;
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

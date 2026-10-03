using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Editorial
{
    public class EditorialGrammarRule : IQualityRule
    {
        public string RuleId => "EDIT_GRAMMAR_PUNCTUATION";
        public string Name => "Repeated Words & Punctuation";
        public string Description => "Detects accidental duplicate words, excessive punctuation marks, and overly lengthy bullet points.";
        public RuleCategory Category => RuleCategory.Editorial;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null) return issues;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;
                if (!SafeInteropHelper.HasText(shape)) continue;

                string text = SafeInteropHelper.GetText(shape);

                // Check repeated words
                var rep = EditorialEngine.CheckRepeatedWords(text);
                foreach (var r in rep)
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
                        Problem = r.Problem,
                        Expected = r.Expected,
                        Actual = r.Actual,
                        Recommendation = r.Recommendation,
                        CanAutoFix = r.CanAutoFix,
                        FixPayload = r.FixedText,
                        BeforeStateSnapshot = text
                    });
                }

                // Check excessive punctuation
                var punct = EditorialEngine.CheckExcessivePunctuation(text);
                foreach (var p in punct)
                {
                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = Name,
                        Category = Category,
                        Severity = IssueSeverity.Info,
                        SlideIndex = context.SlideIndex,
                        ShapeId = shape.Id,
                        ShapeName = shape.Name,
                        Problem = p.Problem,
                        Expected = p.Expected,
                        Actual = p.Actual,
                        Recommendation = p.Recommendation,
                        CanAutoFix = p.CanAutoFix,
                        FixPayload = p.FixedText,
                        BeforeStateSnapshot = text
                    });
                }

                // Check bullet lengths
                var bullets = EditorialEngine.CheckBulletLength(text, 35);
                foreach (var b in bullets)
                {
                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = "Bullet Point Length",
                        Category = Category,
                        Severity = IssueSeverity.Suggestion,
                        SlideIndex = context.SlideIndex,
                        ShapeId = shape.Id,
                        ShapeName = shape.Name,
                        Problem = b.Problem,
                        Expected = b.Expected,
                        Actual = b.Actual,
                        Recommendation = b.Recommendation,
                        CanAutoFix = false
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

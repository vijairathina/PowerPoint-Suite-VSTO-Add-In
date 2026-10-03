using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Alignment
{
    public class OffSlideObjectRule : IQualityRule
    {
        public string RuleId => "LAYOUT_OFF_SLIDE";
        public string Name => "Objects Outside Slide Boundaries";
        public string Description => "Flags shapes or text frames that are placed beyond the visible boundaries of the slide.";
        public RuleCategory Category => RuleCategory.Layout;
        public IssueSeverity DefaultSeverity => IssueSeverity.Error;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => settings?.EnableOffSlideCheck ?? true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null || context.SlideWidth <= 0 || context.SlideHeight <= 0) return issues;

            float w = context.SlideWidth;
            float h = context.SlideHeight;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;

                bool isOff = shape.Left < -5.0f ||
                             shape.Top < -5.0f ||
                             (shape.Left + shape.Width) > (w + 5.0f) ||
                             (shape.Top + shape.Height) > (h + 5.0f);

                if (isOff)
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
                        Problem = $"Shape '{shape.Name}' extends beyond the slide canvas boundary.",
                        Expected = $"Within 0..{w:F0} x 0..{h:F0}",
                        Actual = $"Left:{shape.Left:F0}, Top:{shape.Top:F0}, W:{shape.Width:F0}, H:{shape.Height:F0}",
                        Recommendation = "Reposition or resize shape into safe slide printable boundaries.",
                        CanAutoFix = true,
                        FixPayload = new float[] { w, h }
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
                        float sw = pres.PageSetup.SlideWidth;
                        float sh = pres.PageSetup.SlideHeight;

                        if (s.Left < 20.0f) s.Left = 20.0f;
                        if (s.Top < 20.0f) s.Top = 20.0f;
                        if (s.Left + s.Width > sw - 20.0f) s.Left = sw - s.Width - 20.0f;
                        if (s.Top + s.Height > sh - 20.0f) s.Top = sh - s.Height - 20.0f;

                        issue.IsFixed = true;
                        return true;
                    }
                }
                errorMessage = "Shape not found.";
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

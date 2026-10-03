using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Layout
{
    public class SlideCleanupRule : IQualityRule
    {
        public string RuleId => "LAYOUT_CLEANUP";
        public string Name => "Slide Artifacts & Ghost Shapes";
        public string Description => "Detects empty text boxes, invisible shapes, and zero-dimension artifacts cluttering the slide.";
        public RuleCategory Category => RuleCategory.Layout;
        public IssueSeverity DefaultSeverity => IssueSeverity.Info;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.CurrentSlide == null) return issues;

            var candidates = SlideCleanupService.FindCleanupCandidates(context.CurrentSlide, context.SlideIndex);
            foreach (var c in candidates)
            {
                issues.Add(new QualityIssue
                {
                    RuleId = RuleId,
                    RuleTitle = Name,
                    Category = Category,
                    Severity = DefaultSeverity,
                    SlideIndex = c.SlideIndex,
                    ShapeId = c.ShapeId,
                    ShapeName = c.ShapeName,
                    Problem = c.Reason,
                    Expected = "Clean slide canvas without orphaned or empty objects",
                    Actual = c.Reason,
                    Recommendation = "Remove ghost artifact to reduce presentation bloat and avoid layout confusion.",
                    CanAutoFix = true,
                    FixPayload = c.ShapeId
                });
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
                int deleted = SlideCleanupService.DeleteCandidate(slide, issue.ShapeId);
                if (deleted > 0)
                {
                    issue.IsFixed = true;
                    return true;
                }
                errorMessage = "Could not find artifact on slide.";
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

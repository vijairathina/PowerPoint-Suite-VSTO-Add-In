using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Consistency
{
    public class TitleConsistencyRule : IQualityRule
    {
        public string RuleId => "CONSIST_TITLE_POSITION";
        public string Name => "Cross-Slide Title Consistency";
        public string Description => "Detects title position and font size shifting across presentation slides.";
        public RuleCategory Category => RuleCategory.Consistency;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => true;

        public bool IsEnabled(AssistantSettings settings) => true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.Presentation == null || context.CurrentSlide == null) return issues;

            // Only evaluate during multi-slide scan or compare with Slide 1
            if (context.Presentation.Slides.Count < 2) return issues;

            try
            {
                var slide1 = context.Presentation.Slides[1];
                var refTitle = ConsistencyEngine.DetectSlideTitle(slide1, 1);
                if (refTitle == null || context.SlideIndex == 1) return issues;

                var curTitle = ConsistencyEngine.DetectSlideTitle(context.CurrentSlide, context.SlideIndex);
                if (curTitle == null) return issues;

                float dTop = Math.Abs(curTitle.Top - refTitle.Top);
                float dLeft = Math.Abs(curTitle.Left - refTitle.Left);

                if (dTop > 4.0f || dLeft > 4.0f)
                {
                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = Name,
                        Category = Category,
                        Severity = DefaultSeverity,
                        SlideIndex = context.SlideIndex,
                        ShapeId = curTitle.ShapeId,
                        ShapeName = curTitle.ShapeName,
                        Problem = $"Title position drifts from Slide 1 reference (differs by Top: {dTop:F1}pt, Left: {dLeft:F1}pt).",
                        Expected = $"Top: {refTitle.Top:F1}pt, Left: {refTitle.Left:F1}pt",
                        Actual = $"Top: {curTitle.Top:F1}pt, Left: {curTitle.Left:F1}pt",
                        Recommendation = "Align title to Slide 1 position for consistent visual presentation flow.",
                        CanAutoFix = true,
                        FixPayload = refTitle
                    });
                }
            }
            catch { }

            return issues;
        }

        public bool Fix(QualityIssue issue, PowerPoint.Application app, out string errorMessage)
        {
            errorMessage = string.Empty;
            if (issue == null || app == null) return false;

            try
            {
                var refTitle = issue.FixPayload as SlideTitleMetadata;
                if (refTitle != null)
                {
                    ConsistencyEngine.SyncTitleWithReference(app.ActivePresentation, refTitle, issue.SlideIndex);
                    issue.IsFixed = true;
                    return true;
                }
                errorMessage = "Invalid reference title metadata.";
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

using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Images
{
    public class ImageQualityRule : IQualityRule
    {
        public string RuleId => "IMG_QUALITY_CHECK";
        public string Name => "Image Distortion & Scaling";
        public string Description => "Checks pictures for distorted aspect ratios, excessive upscaling, and oversized asset dimensions.";
        public RuleCategory Category => RuleCategory.Images;
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

                var imgIssues = ImageQualityEngine.InspectShapeImage(shape);
                foreach (var i in imgIssues)
                {
                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = Name,
                        Category = Category,
                        Severity = DefaultSeverity,
                        SlideIndex = context.SlideIndex,
                        ShapeId = i.ShapeId,
                        ShapeName = i.ShapeName,
                        Problem = i.Problem,
                        Expected = "Proper proportions with 1:1 original aspect ratio",
                        Actual = i.Problem,
                        Recommendation = i.Recommendation,
                        CanAutoFix = i.IsDistorted,
                        FixPayload = i.ShapeId
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
                        ImageQualityEngine.ResetAspectRatio(s);
                        issue.IsFixed = true;
                        return true;
                    }
                }
                errorMessage = "Picture shape not found.";
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

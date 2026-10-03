using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Colors
{
    public class ExcessiveColorsRule : IQualityRule
    {
        public string RuleId => "COLOR_EXCESSIVE";
        public string Name => "Excessive Colors on Slide";
        public string Description => "Warns when a slide employs too many distinct colors, undermining visual harmony and brand discipline.";
        public RuleCategory Category => RuleCategory.Colors;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => false;

        public bool IsEnabled(AssistantSettings settings) => settings?.EnableColorContrastCheck ?? true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.CurrentSlide == null) return issues;

            int threshold = context.Settings?.MaxColorsPerSlideThreshold ?? 6;
            var colorDict = ColorEngine.ExtractSlideColors(context.CurrentSlide, context.SlideIndex);

            if (colorDict.Count > threshold)
            {
                string colorSample = string.Join(", ", colorDict.Keys.Take(8));
                issues.Add(new QualityIssue
                {
                    RuleId = RuleId,
                    RuleTitle = Name,
                    Category = Category,
                    Severity = DefaultSeverity,
                    SlideIndex = context.SlideIndex,
                    ShapeId = 0,
                    ShapeName = $"Slide {context.SlideIndex}",
                    Problem = $"Slide contains {colorDict.Count} distinct colors (exceeds recommended max of {threshold}).",
                    Expected = $"<= {threshold} cohesive colors",
                    Actual = $"{colorDict.Count} colors: {colorSample}",
                    Recommendation = "Consolidate shapes to use the approved brand palette to maintain clean visual hierarchy.",
                    CanAutoFix = false
                });
            }

            return issues;
        }

        public bool Fix(QualityIssue issue, PowerPoint.Application app, out string errorMessage)
        {
            errorMessage = "Please use the 'Replace Color' tool in the Color Assistant to map excess colors into your brand palette.";
            return false;
        }
    }
}

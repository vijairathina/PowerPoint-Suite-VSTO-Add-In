using System;
using System.Collections.Generic;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Settings;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Rules.Accessibility
{
    public class MissingAltTextRule : IQualityRule
    {
        public string RuleId => "ACC_MISSING_ALT_TEXT";
        public string Name => "Missing Image Alt Text";
        public string Description => "Checks pictures, diagrams and graphics for missing alternative text descriptions required by accessibility standards.";
        public RuleCategory Category => RuleCategory.Accessibility;
        public IssueSeverity DefaultSeverity => IssueSeverity.Warning;
        public bool CanAutoFix => false;

        public bool IsEnabled(AssistantSettings settings) => true;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null) return issues;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;

                // Pictures, linked pictures, media, or complex grouped diagrams
                bool isVisualAsset = shape.Type == Office.MsoShapeType.msoPicture ||
                                     shape.Type == Office.MsoShapeType.msoLinkedPicture ||
                                     shape.Type == Office.MsoShapeType.msoMedia;

                if (isVisualAsset)
                {
                    string alt = shape.AlternativeText;
                    string title = shape.Title;

                    if (string.IsNullOrWhiteSpace(alt) && string.IsNullOrWhiteSpace(title))
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
                            Problem = $"Image '{shape.Name}' is missing alternative text (Alt Text).",
                            Expected = "Descriptive alternative text for screen readers",
                            Actual = "Empty Alt Text",
                            Recommendation = "Add a concise description of the graphic to support visually impaired viewers.",
                            CanAutoFix = false
                        });
                    }
                }
            }

            return issues;
        }

        public bool Fix(QualityIssue issue, PowerPoint.Application app, out string errorMessage)
        {
            errorMessage = "Please enter an alternative text description in the PowerPoint Format Picture pane.";
            return false;
        }
    }
}

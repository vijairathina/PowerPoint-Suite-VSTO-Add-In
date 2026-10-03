using System.Collections.Generic;
using System.Threading;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Core
{
    public class RuleContext
    {
        public PowerPoint.Application Application { get; set; }
        public PowerPoint.Presentation Presentation { get; set; }
        public PowerPoint.Slide CurrentSlide { get; set; }
        public int SlideIndex { get; set; }
        public List<PowerPoint.Shape> SlideShapes { get; set; } = new List<PowerPoint.Shape>();
        public AnalysisScope Scope { get; set; }
        public AssistantSettings Settings { get; set; }
        public CancellationToken CancellationToken { get; set; }
        public float SlideWidth { get; set; }
        public float SlideHeight { get; set; }
    }

    public interface IQualityRule
    {
        string RuleId { get; }
        string Name { get; }
        string Description { get; }
        RuleCategory Category { get; }
        IssueSeverity DefaultSeverity { get; }
        bool CanAutoFix { get; }

        bool IsEnabled(AssistantSettings settings);
        IEnumerable<QualityIssue> Analyze(RuleContext context);
        bool Fix(QualityIssue issue, PowerPoint.Application app, out string errorMessage);
    }
}

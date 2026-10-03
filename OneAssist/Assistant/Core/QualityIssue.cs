using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Core
{
    /// <summary>
    /// Represents a detected formatting, alignment, color, or editorial issue in PowerPoint.
    /// </summary>
    public class QualityIssue
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string RuleId { get; set; }
        public string RuleTitle { get; set; }
        public RuleCategory Category { get; set; }
        public IssueSeverity Severity { get; set; }

        public int SlideIndex { get; set; }
        public int ShapeId { get; set; }
        public string ShapeName { get; set; }

        public string Description { get; set; }
        public string Problem { get; set; }
        public string Expected { get; set; }
        public string Actual { get; set; }
        public string Recommendation { get; set; }

        public bool CanAutoFix { get; set; }
        public FixSafetyLevel SafetyLevel { get; set; } = FixSafetyLevel.Safe;
        public bool IsFixed { get; set; }
        public bool IsIgnored { get; set; }

        /// <summary>
        /// Additional rule-specific metadata or payload (e.g. proposed color, font, target margin)
        /// </summary>
        public object FixPayload { get; set; }

        /// <summary>
        /// Snapshot of properties prior to fix, allowing rollback.
        /// </summary>
        public object BeforeStateSnapshot { get; set; }

        public override string ToString()
        {
            return $"[{Severity}] Slide {SlideIndex} - {ShapeName}: {Problem}";
        }
    }
}

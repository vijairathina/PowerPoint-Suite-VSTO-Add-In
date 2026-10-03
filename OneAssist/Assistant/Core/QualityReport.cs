using System;
using System.Collections.Generic;
using System.Linq;

namespace OneAssist.Assistant.Core
{
    /// <summary>
    /// Presentation analysis report containing health metrics, issues, and category breakdowns.
    /// </summary>
    public class QualityReport
    {
        public DateTime AnalyzedAt { get; set; } = DateTime.Now;
        public AnalysisScope Scope { get; set; }
        public int TotalSlidesAnalyzed { get; set; }
        public int TotalShapesAnalyzed { get; set; }

        public List<QualityIssue> Issues { get; set; } = new List<QualityIssue>();

        public int TotalIssuesCount => Issues?.Count ?? 0;
        public int ErrorCount => Issues?.Count(i => i.Severity == IssueSeverity.Error && !i.IsIgnored) ?? 0;
        public int WarningCount => Issues?.Count(i => i.Severity == IssueSeverity.Warning && !i.IsIgnored) ?? 0;
        public int InfoCount => Issues?.Count(i => (i.Severity == IssueSeverity.Info || i.Severity == IssueSeverity.Suggestion) && !i.IsIgnored) ?? 0;
        public int FixedCount => Issues?.Count(i => i.IsFixed) ?? 0;
        public int PassedCount => Math.Max(0, (TotalShapesAnalyzed * 3) + (TotalSlidesAnalyzed * 5) - TotalIssuesCount);

        /// <summary>
        /// Overall Presentation Health Score (0 - 100%).
        /// Formula penalizes errors (3 pts) and warnings (1 pt) scaled against shape/slide density.
        /// </summary>
        public double HealthScore
        {
            get
            {
                if (TotalShapesAnalyzed == 0 && TotalSlidesAnalyzed == 0) return 100.0;
                double penalty = (ErrorCount * 3.5) + (WarningCount * 1.2) + (InfoCount * 0.3);
                double basePoints = Math.Max(10.0, (TotalShapesAnalyzed * 2.0) + (TotalSlidesAnalyzed * 5.0));
                double score = Math.Max(0.0, 100.0 - (penalty / basePoints * 100.0));
                return Math.Round(score, 1);
            }
        }

        public double GetCategoryScore(RuleCategory category)
        {
            var catIssues = Issues?.Where(i => i.Category == category && !i.IsIgnored).ToList();
            if (catIssues == null || catIssues.Count == 0) return 100.0;

            int catErrors = catIssues.Count(i => i.Severity == IssueSeverity.Error);
            int catWarnings = catIssues.Count(i => i.Severity == IssueSeverity.Warning);
            double penalty = (catErrors * 4.0) + (catWarnings * 1.5);
            double basePoints = Math.Max(5.0, TotalShapesAnalyzed * 1.5);
            return Math.Round(Math.Max(0.0, 100.0 - (penalty / basePoints * 100.0)), 1);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Core;

namespace OneAssist.Assistant.Scorecard
{
    public class CategoryScorecardItem
    {
        public string CategoryName { get; set; }
        public double Score { get; set; }
        public int ErrorCount { get; set; }
        public int WarningCount { get; set; }
        public string StatusBadge => Score >= 90 ? "Passed" : (Score >= 70 ? "Warning" : "Attention");
    }

    public static class ScorecardService
    {
        public static List<CategoryScorecardItem> GenerateScorecard(QualityReport report)
        {
            var scorecard = new List<CategoryScorecardItem>();
            if (report == null) return scorecard;

            var categories = new[]
            {
                RuleCategory.Formatting,
                RuleCategory.Typography,
                RuleCategory.Alignment,
                RuleCategory.Colors,
                RuleCategory.Layout,
                RuleCategory.Editorial,
                RuleCategory.Consistency,
                RuleCategory.Accessibility,
                RuleCategory.Images,
                RuleCategory.Branding
            };

            foreach (var cat in categories)
            {
                var issues = report.Issues.Where(i => i.Category == cat && !i.IsIgnored).ToList();
                int errors = issues.Count(i => i.Severity == IssueSeverity.Error);
                int warnings = issues.Count(i => i.Severity == IssueSeverity.Warning);
                double score = report.GetCategoryScore(cat);

                scorecard.Add(new CategoryScorecardItem
                {
                    CategoryName = cat.ToString(),
                    Score = score,
                    ErrorCount = errors,
                    WarningCount = warnings
                });
            }

            return scorecard;
        }
    }
}

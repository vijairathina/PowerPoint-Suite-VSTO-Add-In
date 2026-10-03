using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Core
{
    public class FixResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public int FixedCount { get; set; }
    }

    public class FixEngine
    {
        private readonly RuleRegistry _ruleRegistry;

        public FixEngine(RuleRegistry registry = null)
        {
            _ruleRegistry = registry ?? RuleRegistry.Instance;
        }

        /// <summary>
        /// Navigates PowerPoint view to the slide and highlights the shape with the issue.
        /// </summary>
        public bool NavigateToIssue(PowerPoint.Application app, QualityIssue issue)
        {
            if (app == null || issue == null) return false;
            try
            {
                var pres = app.ActivePresentation;
                if (pres == null) return false;

                var slide = pres.Slides[issue.SlideIndex];
                slide.Select();

                if (issue.ShapeId > 0)
                {
                    foreach (PowerPoint.Shape s in slide.Shapes)
                    {
                        if (s.Id == issue.ShapeId)
                        {
                            s.Select();
                            return true;
                        }
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Fixes a specific issue using its registered rule.
        /// </summary>
        public FixResult FixIssue(PowerPoint.Application app, QualityIssue issue)
        {
            if (app == null || issue == null)
            {
                return new FixResult { Success = false, Message = "Invalid issue or application reference." };
            }

            if (issue.IsFixed)
            {
                return new FixResult { Success = true, Message = "Issue is already marked as fixed." };
            }

            var rule = _ruleRegistry.GetRuleById(issue.RuleId);
            if (rule == null)
            {
                return new FixResult { Success = false, Message = $"Rule '{issue.RuleId}' is not registered." };
            }

            bool ok = rule.Fix(issue, app, out string error);
            return new FixResult
            {
                Success = ok,
                Message = ok ? "Fix applied successfully." : error,
                FixedCount = ok ? 1 : 0
            };
        }

        /// <summary>
        /// Applies fixes for all auto-fixable issues in the list.
        /// </summary>
        public FixResult FixAll(PowerPoint.Application app, IEnumerable<QualityIssue> issues)
        {
            if (app == null || issues == null)
            {
                return new FixResult { Success = false, Message = "Invalid parameters." };
            }

            int fixedCount = 0;
            var list = issues.Where(i => i.CanAutoFix && !i.IsFixed && !i.IsIgnored).ToList();

            foreach (var issue in list)
            {
                var res = FixIssue(app, issue);
                if (res.Success)
                {
                    fixedCount++;
                }
            }

            return new FixResult
            {
                Success = fixedCount > 0,
                Message = $"Successfully fixed {fixedCount} of {list.Count} issues.",
                FixedCount = fixedCount
            };
        }
    }
}

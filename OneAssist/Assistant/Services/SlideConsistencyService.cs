using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public class ConsistencyDifferenceReport
    {
        public int SlideIndex { get; set; }
        public List<QualityIssue> TypographyDifferences { get; set; } = new List<QualityIssue>();
        public List<QualityIssue> ColorDifferences { get; set; } = new List<QualityIssue>();
        public List<QualityIssue> AlignmentDifferences { get; set; } = new List<QualityIssue>();
        public List<QualityIssue> SpacingDifferences { get; set; } = new List<QualityIssue>();
        public List<QualityIssue> BrandingDifferences { get; set; } = new List<QualityIssue>();

        public List<QualityIssue> AllDifferences
        {
            get
            {
                var list = new List<QualityIssue>();
                list.AddRange(TypographyDifferences);
                list.AddRange(ColorDifferences);
                list.AddRange(AlignmentDifferences);
                list.AddRange(SpacingDifferences);
                list.AddRange(BrandingDifferences);
                return list;
            }
        }
    }

    public static class SlideConsistencyService
    {
        /// <summary>
        /// Compares the current slide against reference slide, brand rules, and presentation patterns (Req 47).
        /// </summary>
        public static ConsistencyDifferenceReport AnalyzeSlideConsistency(PowerPoint.Slide currentSlide, PowerPoint.Presentation pres, AssistantSettings settings)
        {
            var diff = new ConsistencyDifferenceReport();
            if (currentSlide == null || pres == null) return diff;

            diff.SlideIndex = currentSlide.SlideIndex;

            try
            {
                var shapes = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape s in currentSlide.Shapes) shapes.Add(s);

                var context = new RuleContext
                {
                    Application = currentSlide.Application,
                    Presentation = pres,
                    CurrentSlide = currentSlide,
                    SlideIndex = currentSlide.SlideIndex,
                    SlideShapes = shapes,
                    Scope = AnalysisScope.ActiveSlide,
                    Settings = settings,
                    SlideWidth = pres.PageSetup.SlideWidth,
                    SlideHeight = pres.PageSetup.SlideHeight
                };

                // Evaluate all registered rules
                var rules = RuleRegistry.Instance.GetAllRules();
                foreach (var rule in rules)
                {
                    try
                    {
                        var issues = rule.Analyze(context);
                        if (issues != null)
                        {
                            foreach (var iss in issues)
                            {
                                switch (iss.Category)
                                {
                                    case RuleCategory.Typography:
                                    case RuleCategory.Formatting:
                                        diff.TypographyDifferences.Add(iss);
                                        break;
                                    case RuleCategory.Colors:
                                        diff.ColorDifferences.Add(iss);
                                        break;
                                    case RuleCategory.Alignment:
                                    case RuleCategory.Layout:
                                        if (iss.RuleId.Contains("SPACING"))
                                            diff.SpacingDifferences.Add(iss);
                                        else
                                            diff.AlignmentDifferences.Add(iss);
                                        break;
                                    case RuleCategory.Branding:
                                    case RuleCategory.Consistency:
                                        diff.BrandingDifferences.Add(iss);
                                        break;
                                    default:
                                        diff.TypographyDifferences.Add(iss);
                                        break;
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return diff;
        }

        /// <summary>
        /// Analyzes the entire presentation and organizes consistency violations by slide (Req 48).
        /// </summary>
        public static Dictionary<int, List<QualityIssue>> AnalyzePresentationConsistency(PowerPoint.Presentation pres, AssistantSettings settings)
        {
            var deckIssues = new Dictionary<int, List<QualityIssue>>();
            if (pres == null) return deckIssues;

            try
            {
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    var slide = pres.Slides[i];
                    var diff = AnalyzeSlideConsistency(slide, pres, settings);
                    if (diff.AllDifferences.Count > 0)
                    {
                        deckIssues[i] = diff.AllDifferences;
                    }
                }
            }
            catch { }

            return deckIssues;
        }

        /// <summary>
        /// Fixes all issues matching the requested safety level (Safe only by default).
        /// </summary>
        public static int FixSafeIssues(PowerPoint.Application app, IEnumerable<QualityIssue> issues, FixSafetyLevel maxSafety = FixSafetyLevel.Safe)
        {
            if (app == null || issues == null) return 0;
            int fixedCount = 0;
            var fixEngine = new FixEngine();

            foreach (var issue in issues.Where(i => i.CanAutoFix && !i.IsFixed && i.SafetyLevel <= maxSafety))
            {
                var res = fixEngine.FixIssue(app, issue);
                if (res.Success)
                {
                    fixedCount++;
                }
            }

            return fixedCount;
        }
    }
}

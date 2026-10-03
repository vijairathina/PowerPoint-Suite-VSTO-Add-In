using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OneAssist.Assistant.AILayout.Models;
using OneAssist.Assistant.AILayout.Templates;
using OneAssist.Assistant.Settings;

namespace OneAssist.Assistant.AILayout.Planner
{
    public class LocalHeuristicLayoutPlanner : IAILayoutPlanner
    {
        public AIProviderType ProviderType => AIProviderType.LocalHeuristic;

        public Task<List<LayoutSuggestion>> PlanLayoutsAsync(SlideSemanticAnalysis analysis, AssistantSettings settings, DesignStyle style, CancellationToken ct)
        {
            var suggestions = new List<LayoutSuggestion>();
            if (analysis == null) return Task.FromResult(suggestions);

            int optNum = 1;

            switch (analysis.DetectedContentType)
            {
                case SlideContentType.KPI:
                case SlideContentType.Chart:
                    // 1. KPI + Chart
                    suggestions.Add(CreateSuggestion(optNum++, "KPI + Supporting Chart",
                        LayoutTemplatePattern.KPIPlusChart,
                        "Recommended because the slide features headline numerical metrics backed by supporting data charts. Separating the metric and visualization provides immediate executive clarity.",
                        LayoutTemplateCatalog.CreateKPIPlusChartLayout(analysis), "📊", "High Data Impact", "Optimal"));

                    // 2. Executive Summary
                    suggestions.Add(CreateSuggestion(optNum++, "Executive Summary",
                        LayoutTemplatePattern.ExecutiveSummary,
                        "Frames the headline finding at the top with supporting regional/metric pillars underneath.",
                        LayoutTemplateCatalog.CreateExecutiveSummaryLayout(analysis), "👔", "Executive Cadence", "Balanced"));

                    // 3. Three Cards
                    suggestions.Add(CreateSuggestion(optNum++, "Three Metric Cards",
                        LayoutTemplatePattern.ThreeCards,
                        "Distributes metrics and supporting commentary across three equal-width cards for structured reading.",
                        LayoutTemplateCatalog.CreateThreeCardsLayout(analysis), "🃏", "Balanced Scanability", "Balanced"));

                    // 4. Minimal Focus
                    suggestions.Add(CreateSuggestion(optNum++, "Minimal Data Focus",
                        LayoutTemplatePattern.OneBigNumber,
                        "Maximizes visual whitespace and focuses executive attention solely on the primary milestone.",
                        LayoutTemplateCatalog.CreateMinimalFocusLayout(analysis), "🎯", "Maximum Focus", "Light"));
                    break;

                case SlideContentType.Comparison:
                    // 1. Two Column Comparison
                    suggestions.Add(CreateSuggestion(optNum++, "Two-Column Comparison",
                        LayoutTemplatePattern.TwoColumnComparison,
                        "Aligns contrasting concepts, options, or before/after items into balanced side-by-side columns.",
                        LayoutTemplateCatalog.CreateTwoColumnComparisonLayout(analysis), "⚖️", "Direct Contrast", "Optimal"));

                    // 2. Three Cards
                    suggestions.Add(CreateSuggestion(optNum++, "Multi-Option Pillars",
                        LayoutTemplatePattern.ThreeCards,
                        "Organizes multi-option benchmarking across three distinct cards.",
                        LayoutTemplateCatalog.CreateThreeCardsLayout(analysis), "🃏", "Structured Comparison", "Balanced"));

                    // 3. Executive Summary
                    suggestions.Add(CreateSuggestion(optNum++, "Executive Takeaway",
                        LayoutTemplatePattern.ExecutiveSummary,
                        "Highlights the comparative conclusion at the top with granular option details below.",
                        LayoutTemplateCatalog.CreateExecutiveSummaryLayout(analysis), "👔", "Executive Clarity", "Balanced"));
                    break;

                case SlideContentType.ImageHeavy:
                    // 1. Text + Image
                    suggestions.Add(CreateSuggestion(optNum++, "Visual Hero Story",
                        LayoutTemplatePattern.TextPlusImage,
                        "Balances the hero image on the left with structured key takeaway copy on the right.",
                        LayoutTemplateCatalog.CreateTextPlusImageLayout(analysis), "🖼️", "High Visual Engagement", "Balanced"));

                    // 2. Executive Summary
                    suggestions.Add(CreateSuggestion(optNum++, "Visual Narrative + Pillars",
                        LayoutTemplatePattern.ExecutiveSummary,
                        "Elevates the core message with an executive takeaway banner and visual evidence cards.",
                        LayoutTemplateCatalog.CreateExecutiveSummaryLayout(analysis), "👔", "Executive Impact", "Balanced"));

                    // 3. Minimal Focus
                    suggestions.Add(CreateSuggestion(optNum++, "Minimalist Showcase",
                        LayoutTemplatePattern.OneBigNumber,
                        "Spacious gallery presentation with generous margins and clean typography.",
                        LayoutTemplateCatalog.CreateMinimalFocusLayout(analysis), "✨", "Refined Polish", "Light"));
                    break;

                case SlideContentType.TextHeavy:
                default:
                    // 1. Three Cards
                    suggestions.Add(CreateSuggestion(optNum++, "Three Structured Cards",
                        LayoutTemplatePattern.ThreeCards,
                        $"Restructures {analysis.TotalWordCount} words and {analysis.TotalBulletCount} bullet points into 3 digestible visual containers, removing list clutter.",
                        LayoutTemplateCatalog.CreateThreeCardsLayout(analysis), "🃏", "Clutter Elimination", "Balanced"));

                    // 2. Two Column
                    suggestions.Add(CreateSuggestion(optNum++, "Two-Column Split",
                        LayoutTemplatePattern.TwoColumnComparison,
                        "Divides dense narrative text into two balanced reading columns for improved readability.",
                        LayoutTemplateCatalog.CreateTwoColumnComparisonLayout(analysis), "📑", "Readable Hierarchy", "Balanced"));

                    // 3. Executive Summary
                    suggestions.Add(CreateSuggestion(optNum++, "Executive Summary",
                        LayoutTemplatePattern.ExecutiveSummary,
                        "Elevates the single most important conclusion to a headline box and groups the remaining text as evidence.",
                        LayoutTemplateCatalog.CreateExecutiveSummaryLayout(analysis), "👔", "Executive Cadence", "Optimal"));

                    // 4. Minimal Focus
                    suggestions.Add(CreateSuggestion(optNum++, "Key Message Minimal",
                        LayoutTemplatePattern.OneBigNumber,
                        "Highlights the central thesis with prominent headline typography and intentional white space.",
                        LayoutTemplateCatalog.CreateMinimalFocusLayout(analysis), "🎯", "Maximum Focus", "Light"));
                    break;
            }

            return Task.FromResult(suggestions);
        }

        private static LayoutSuggestion CreateSuggestion(int number, string title, LayoutTemplatePattern pattern, string reasoning, LayoutSpecification spec, string icon, string impact, string density)
        {
            return new LayoutSuggestion
            {
                OptionNumber = number,
                Title = title,
                Pattern = pattern,
                Reasoning = reasoning,
                ObjectCount = spec.Regions.Count,
                ContentDensity = density,
                DesignImpact = impact,
                IconSymbol = icon,
                QualityScore = new LayoutQualityScore
                {
                    Hierarchy = "Excellent",
                    Spacing = "Excellent",
                    Readability = "Good",
                    BrandConsistency = "Excellent",
                    ContentPreservationPct = 100.0,
                    Complexity = "Low",
                    Confidence = "Strong structural match"
                },
                Specification = spec
            };
        }
    }
}

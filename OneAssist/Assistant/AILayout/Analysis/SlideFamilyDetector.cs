using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.AILayout.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.AILayout.Analysis
{
    public class SlideFamily
    {
        public string FamilyName { get; set; } = string.Empty;
        public SlideContentType ContentType { get; set; }
        public string Description { get; set; } = string.Empty;
        public List<int> SlideIndices { get; set; } = new List<int>();
    }

    public static class SlideFamilyDetector
    {
        public static List<SlideFamily> DetectFamilies(PowerPoint.Presentation pres)
        {
            var families = new List<SlideFamily>();
            if (pres == null || pres.Slides.Count == 0) return families;

            var familyMap = new Dictionary<SlideContentType, List<int>>();

            for (int i = 1; i <= pres.Slides.Count; i++)
            {
                try
                {
                    var slide = pres.Slides[i];
                    var analysis = SlideSemanticAnalyzer.Analyze(slide);
                    if (!familyMap.ContainsKey(analysis.DetectedContentType))
                    {
                        familyMap[analysis.DetectedContentType] = new List<int>();
                    }
                    familyMap[analysis.DetectedContentType].Add(i);
                }
                catch { }
            }

            foreach (var kvp in familyMap.OrderByDescending(k => k.Value.Count))
            {
                families.Add(new SlideFamily
                {
                    ContentType = kvp.Key,
                    FamilyName = GetFamilyDisplayName(kvp.Key),
                    Description = $"{kvp.Value.Count} slide(s) sharing {kvp.Key} structure",
                    SlideIndices = kvp.Value
                });
            }

            return families;
        }

        private static string GetFamilyDisplayName(SlideContentType type)
        {
            switch (type)
            {
                case SlideContentType.ExecutiveSummary: return "Executive Summary Slides";
                case SlideContentType.KPI: return "KPI & Performance Metrics";
                case SlideContentType.Chart: return "Data & Chart Slides";
                case SlideContentType.Table: return "Tabular Data Slides";
                case SlideContentType.Comparison: return "Comparative & Benchmarking Slides";
                case SlideContentType.Timeline:
                case SlideContentType.Process: return "Process & Roadmap Slides";
                case SlideContentType.Strategy: return "Strategy & Framework Slides";
                case SlideContentType.ImageHeavy: return "Visual & Media Gallery";
                case SlideContentType.TextHeavy: return "Detailed Content & Narrative";
                case SlideContentType.SectionDivider: return "Section Dividers";
                default: return "General Topic Slides";
            }
        }
    }
}

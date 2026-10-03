using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.AILayout.Analysis;
using OneAssist.Assistant.AILayout.Models;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.AILayout.Engine
{
    public class SequenceAnalysisResult
    {
        public int TotalSlidesAnalyzed { get; set; }
        public SlideContentType DominantFamily { get; set; }
        public List<string> SequenceInsights { get; set; } = new List<string>();
        public int HarmonizedCount { get; set; }
    }

    public static class SlideSequenceEngine
    {
        public static SequenceAnalysisResult AnalyzeSequence(PowerPoint.Presentation pres, IEnumerable<int> slideIndices)
        {
            var res = new SequenceAnalysisResult();
            if (pres == null || slideIndices == null) return res;

            var analyses = new List<SlideSemanticAnalysis>();
            foreach (int idx in slideIndices)
            {
                if (idx >= 1 && idx <= pres.Slides.Count)
                {
                    try
                    {
                        var slide = pres.Slides[idx];
                        analyses.Add(SlideSemanticAnalyzer.Analyze(slide));
                    }
                    catch { }
                }
            }

            res.TotalSlidesAnalyzed = analyses.Count;
            if (analyses.Count == 0) return res;

            // Find dominant content type
            var grouped = new Dictionary<SlideContentType, int>();
            foreach (var a in analyses)
            {
                if (!grouped.ContainsKey(a.DetectedContentType)) grouped[a.DetectedContentType] = 0;
                grouped[a.DetectedContentType]++;
            }

            var top = System.Linq.Enumerable.OrderByDescending(grouped, g => g.Value).First();
            res.DominantFamily = top.Key;
            res.SequenceInsights.Add($"Dominant pattern detected across sequence: {top.Key} ({top.Value}/{analyses.Count} slides).");

            return res;
        }

        public static int HarmonizeSlides(PowerPoint.Presentation pres, IEnumerable<int> slideIndices, AssistantSettings settings)
        {
            if (pres == null || slideIndices == null) return 0;
            int count = 0;

            var indicesList = new List<int>(slideIndices);
            if (indicesList.Count < 2) return 0;

            try
            {
                // Use first slide as visual benchmark
                var benchmarkSlide = pres.Slides[indicesList[0]];

                for (int i = 1; i < indicesList.Count; i++)
                {
                    var targetSlide = pres.Slides[indicesList[i]];
                    var spec = ReferenceSlideEngine.AdaptReferenceSlideLayout(benchmarkSlide, targetSlide);
                    if (spec != null)
                    {
                        bool success = PowerPointLayoutRenderer.ApplyLayout(targetSlide, spec, settings, false);
                        if (success) count++;
                    }
                }
            }
            catch { }

            return count;
        }
    }
}

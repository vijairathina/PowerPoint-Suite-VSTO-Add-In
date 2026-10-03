using System;
using System.Collections.Generic;

namespace OneAssist.Assistant.Core
{
    public class CachedSlideAnalysis
    {
        public int SlideIndex { get; set; }
        public int ShapesCount { get; set; }
        public DateTime CachedAt { get; set; } = DateTime.Now;
        public List<QualityIssue> Issues { get; set; } = new List<QualityIssue>();
    }

    public static class AnalysisCacheManager
    {
        private static readonly Dictionary<int, CachedSlideAnalysis> _cache = new Dictionary<int, CachedSlideAnalysis>();

        public static bool TryGetCached(int slideIndex, int currentShapesCount, out List<QualityIssue> issues)
        {
            issues = null;
            if (_cache.TryGetValue(slideIndex, out var cached))
            {
                // If shapes count hasn't changed and cache is within 5 minutes
                if (cached.ShapesCount == currentShapesCount && (DateTime.Now - cached.CachedAt).TotalMinutes < 5)
                {
                    issues = new List<QualityIssue>(cached.Issues);
                    return true;
                }
            }
            return false;
        }

        public static void SetCache(int slideIndex, int shapesCount, List<QualityIssue> issues)
        {
            _cache[slideIndex] = new CachedSlideAnalysis
            {
                SlideIndex = slideIndex,
                ShapesCount = shapesCount,
                Issues = new List<QualityIssue>(issues ?? new List<QualityIssue>()),
                CachedAt = DateTime.Now
            };
        }

        public static void InvalidateAll()
        {
            _cache.Clear();
        }

        public static void InvalidateSlide(int slideIndex)
        {
            _cache.Remove(slideIndex);
        }
    }
}

using System;
using System.Collections.Generic;

namespace OneAssist.Assistant.Services
{
    public class LocalUsageStats
    {
        public int TotalPresentationsScanned { get; set; }
        public int TotalFixesApplied { get; set; }
        public int BrandViolationsFixed { get; set; }
        public Dictionary<string, int> ToolUsageFrequency { get; set; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    }

    public static class LocalAnalyticsService
    {
        private static readonly LocalUsageStats _stats = new LocalUsageStats();

        public static LocalUsageStats CurrentStats => _stats;

        public static void RecordToolUsage(string toolName)
        {
            if (string.IsNullOrWhiteSpace(toolName)) return;

            if (!_stats.ToolUsageFrequency.ContainsKey(toolName))
            {
                _stats.ToolUsageFrequency[toolName] = 0;
            }
            _stats.ToolUsageFrequency[toolName]++;
        }

        public static void RecordFixApplied(bool isBrandFix = false)
        {
            _stats.TotalFixesApplied++;
            if (isBrandFix) _stats.BrandViolationsFixed++;
        }

        public static void RecordScanCompleted()
        {
            _stats.TotalPresentationsScanned++;
        }
    }
}

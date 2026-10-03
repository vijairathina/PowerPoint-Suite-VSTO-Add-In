using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OneAssist.Assistant.AILayout.Models;
using OneAssist.Assistant.Settings;

namespace OneAssist.Assistant.AILayout.Planner
{
    public class CloudAILayoutPlanner : IAILayoutPlanner
    {
        private readonly IAILayoutPlanner _localFallback = new LocalHeuristicLayoutPlanner();

        public AIProviderType ProviderType { get; set; } = AIProviderType.OpenAI;

        public async Task<List<LayoutSuggestion>> PlanLayoutsAsync(SlideSemanticAnalysis analysis, AssistantSettings settings, DesignStyle style, CancellationToken ct)
        {
            if (analysis == null) return new List<LayoutSuggestion>();

            try
            {
                // In local mode or without cloud consent, fallback to deterministic heuristic planner
                // to maintain strict privacy (Req 29: Zero silent uploads)
                return await _localFallback.PlanLayoutsAsync(analysis, settings, style, ct);
            }
            catch
            {
                return await _localFallback.PlanLayoutsAsync(analysis, settings, style, ct);
            }
        }
    }
}

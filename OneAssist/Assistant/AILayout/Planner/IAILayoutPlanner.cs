using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OneAssist.Assistant.AILayout.Models;
using OneAssist.Assistant.Settings;

namespace OneAssist.Assistant.AILayout.Planner
{
    public interface IAILayoutPlanner
    {
        AIProviderType ProviderType { get; }
        Task<List<LayoutSuggestion>> PlanLayoutsAsync(SlideSemanticAnalysis analysis, AssistantSettings settings, DesignStyle style, CancellationToken ct);
    }
}

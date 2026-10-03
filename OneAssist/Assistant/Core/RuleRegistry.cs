using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Rules.Accessibility;
using OneAssist.Assistant.Rules.Alignment;
using OneAssist.Assistant.Rules.Colors;
using OneAssist.Assistant.Rules.Consistency;
using OneAssist.Assistant.Rules.Editorial;
using OneAssist.Assistant.Rules.Formatting;
using OneAssist.Assistant.Rules.Images;
using OneAssist.Assistant.Rules.Layout;

namespace OneAssist.Assistant.Core
{
    public class RuleRegistry
    {
        private static RuleRegistry _instance;
        public static RuleRegistry Instance => _instance ?? (_instance = new RuleRegistry());

        private readonly List<IQualityRule> _rules = new List<IQualityRule>();

        public RuleRegistry()
        {
            RegisterDefaultRules();
        }

        private void RegisterDefaultRules()
        {
            // Formatting & Typography
            RegisterRule(new FontConsistencyRule());
            RegisterRule(new FontSizeConsistencyRule());
            RegisterRule(new TextBoxMarginRule());

            // Alignment & Layout
            RegisterRule(new AlmostAlignedRule());
            RegisterRule(new InconsistentSpacingRule());
            RegisterRule(new OffSlideObjectRule());
            RegisterRule(new SlideCleanupRule());

            // Colors
            RegisterRule(new ColorContrastRule());
            RegisterRule(new ExcessiveColorsRule());

            // Editorial
            RegisterRule(new EditorialTerminologyRule());
            RegisterRule(new EditorialGrammarRule());

            // Consistency
            RegisterRule(new TitleConsistencyRule());

            // Accessibility
            RegisterRule(new MissingAltTextRule());

            // Images
            RegisterRule(new ImageQualityRule());
        }

        public void RegisterRule(IQualityRule rule)
        {
            if (rule == null) throw new ArgumentNullException(nameof(rule));
            if (!_rules.Any(r => r.RuleId.Equals(rule.RuleId, StringComparison.OrdinalIgnoreCase)))
            {
                _rules.Add(rule);
            }
        }

        public IReadOnlyList<IQualityRule> GetAllRules() => _rules.AsReadOnly();

        public IEnumerable<IQualityRule> GetRulesByCategory(RuleCategory category)
        {
            return _rules.Where(r => r.Category == category);
        }

        public IQualityRule GetRuleById(string ruleId)
        {
            return _rules.FirstOrDefault(r => r.RuleId.Equals(ruleId, StringComparison.OrdinalIgnoreCase));
        }
    }
}

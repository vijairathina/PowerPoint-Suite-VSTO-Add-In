using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Core
{
    [DataContract]
    public class CustomRuleDefinition
    {
        [DataMember] public string Id { get; set; } = Guid.NewGuid().ToString("N");
        [DataMember] public string Name { get; set; }
        [DataMember] public string TargetProperty { get; set; } // "Font.Name", "Font.Size", "Shape.Width", "Shape.Height"
        [DataMember] public string Operator { get; set; }       // "==", "!=", "<", ">", "Contains"
        [DataMember] public string TargetValue { get; set; }
        [DataMember] public IssueSeverity Severity { get; set; } = IssueSeverity.Warning;
        [DataMember] public string Message { get; set; }
        [DataMember] public bool IsEnabled { get; set; } = true;
    }

    public class DynamicCustomRule : IQualityRule
    {
        private readonly CustomRuleDefinition _def;

        public DynamicCustomRule(CustomRuleDefinition def)
        {
            _def = def ?? throw new ArgumentNullException(nameof(def));
        }

        public string RuleId => $"CUSTOM_{_def.Id}";
        public string Name => _def.Name ?? "Custom Rule";
        public string Description => _def.Message ?? "Custom user-defined validation rule.";
        public RuleCategory Category => RuleCategory.Formatting;
        public IssueSeverity DefaultSeverity => _def.Severity;
        public bool CanAutoFix => false;

        public bool IsEnabled(AssistantSettings settings) => _def.IsEnabled;

        public IEnumerable<QualityIssue> Analyze(RuleContext context)
        {
            var issues = new List<QualityIssue>();
            if (context?.SlideShapes == null || string.IsNullOrWhiteSpace(_def.TargetProperty)) return issues;

            foreach (var shape in context.SlideShapes)
            {
                if (context.CancellationToken.IsCancellationRequested) break;

                bool matches = false;
                string actualValue = string.Empty;

                if (_def.TargetProperty.Equals("Font.Name", StringComparison.OrdinalIgnoreCase))
                {
                    if (SafeInteropHelper.HasText(shape))
                    {
                        actualValue = SafeInteropHelper.GetFontName(shape) ?? string.Empty;
                        matches = EvaluateString(actualValue, _def.Operator, _def.TargetValue);
                    }
                }
                else if (_def.TargetProperty.Equals("Font.Size", StringComparison.OrdinalIgnoreCase))
                {
                    if (SafeInteropHelper.HasText(shape))
                    {
                        float? sz = SafeInteropHelper.GetFontSize(shape);
                        if (sz.HasValue)
                        {
                            actualValue = sz.Value.ToString("F1");
                            matches = EvaluateNumeric(sz.Value, _def.Operator, _def.TargetValue);
                        }
                    }
                }
                else if (_def.TargetProperty.Equals("Shape.Width", StringComparison.OrdinalIgnoreCase))
                {
                    actualValue = shape.Width.ToString("F0");
                    matches = EvaluateNumeric(shape.Width, _def.Operator, _def.TargetValue);
                }

                if (matches)
                {
                    issues.Add(new QualityIssue
                    {
                        RuleId = RuleId,
                        RuleTitle = Name,
                        Category = Category,
                        Severity = DefaultSeverity,
                        SlideIndex = context.SlideIndex,
                        ShapeId = shape.Id,
                        ShapeName = shape.Name,
                        Problem = _def.Message ?? $"Custom rule violation: {actualValue} {_def.Operator} {_def.TargetValue}",
                        Expected = $"{_def.TargetProperty} {_def.Operator} {_def.TargetValue}",
                        Actual = actualValue,
                        Recommendation = "Adjust property to satisfy custom corporate policy rule.",
                        CanAutoFix = false
                    });
                }
            }

            return issues;
        }

        private bool EvaluateString(string actual, string op, string target)
        {
            if (op == "==") return string.Equals(actual, target, StringComparison.OrdinalIgnoreCase);
            if (op == "!=") return !string.Equals(actual, target, StringComparison.OrdinalIgnoreCase);
            if (op == "Contains") return actual.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0;
            return false;
        }

        private bool EvaluateNumeric(float actual, string op, string targetStr)
        {
            if (float.TryParse(targetStr, out float target))
            {
                if (op == "<") return actual < target;
                if (op == ">") return actual > target;
                if (op == "==") return Math.Abs(actual - target) < 0.1f;
                if (op == "!=") return Math.Abs(actual - target) >= 0.1f;
            }
            return false;
        }

        public bool Fix(QualityIssue issue, PowerPoint.Application app, out string errorMessage)
        {
            errorMessage = "Custom rules require manual adjustment.";
            return false;
        }
    }
}

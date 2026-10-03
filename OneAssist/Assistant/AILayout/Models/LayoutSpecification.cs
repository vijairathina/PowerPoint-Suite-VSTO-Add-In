using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace OneAssist.Assistant.AILayout.Models
{
    [DataContract]
    public class LayoutRegion
    {
        [DataMember(Name = "name")]
        public string Name { get; set; } = string.Empty;

        [DataMember(Name = "x")]
        public float X { get; set; } // Left position in points

        [DataMember(Name = "y")]
        public float Y { get; set; } // Top position in points

        [DataMember(Name = "width")]
        public float Width { get; set; }

        [DataMember(Name = "height")]
        public float Height { get; set; }

        [DataMember(Name = "preferred_type")]
        public string PreferredType { get; set; } = string.Empty;

        [DataMember(Name = "text_align")]
        public string TextAlign { get; set; } = "Left";

        [DataMember(Name = "is_container")]
        public bool IsContainer { get; set; }

        [DataMember(Name = "background_color")]
        public string BackgroundColorHex { get; set; }

        [DataMember(Name = "border_color")]
        public string BorderColorHex { get; set; }

        [DataMember(Name = "corner_radius")]
        public float CornerRadius { get; set; } = 4f;
    }

    [DataContract]
    public class ObjectAssignment
    {
        [DataMember(Name = "shape_id")]
        public int ShapeId { get; set; }

        [DataMember(Name = "target_region")]
        public string TargetRegion { get; set; } = string.Empty;

        [DataMember(Name = "role")]
        public string Role { get; set; } = string.Empty;

        [DataMember(Name = "preserve_formatting")]
        public bool PreserveFormatting { get; set; } = true;
    }

    [DataContract]
    public class LayoutQualityScore
    {
        [DataMember(Name = "hierarchy")]
        public string Hierarchy { get; set; } = "Excellent";

        [DataMember(Name = "spacing")]
        public string Spacing { get; set; } = "Excellent";

        [DataMember(Name = "readability")]
        public string Readability { get; set; } = "Good";

        [DataMember(Name = "brand_consistency")]
        public string BrandConsistency { get; set; } = "Excellent";

        [DataMember(Name = "content_preservation_pct")]
        public double ContentPreservationPct { get; set; } = 100.0;

        [DataMember(Name = "complexity")]
        public string Complexity { get; set; } = "Low";

        [DataMember(Name = "confidence")]
        public string Confidence { get; set; } = "Strong structural match";
    }

    [DataContract]
    public class LayoutSpecification
    {
        [DataMember(Name = "slide_type")]
        public string SlideType { get; set; } = "general_content";

        [DataMember(Name = "layout")]
        public string Layout { get; set; } = "two_column";

        [DataMember(Name = "regions")]
        public List<LayoutRegion> Regions { get; set; } = new List<LayoutRegion>();

        [DataMember(Name = "object_assignments")]
        public List<ObjectAssignment> ObjectAssignments { get; set; } = new List<ObjectAssignment>();

        [DataMember(Name = "reasoning_summary")]
        public string ReasoningSummary { get; set; } = string.Empty;

        [DataMember(Name = "warnings")]
        public List<string> Warnings { get; set; } = new List<string>();

        [DataMember(Name = "quality_score")]
        public LayoutQualityScore QualityScore { get; set; } = new LayoutQualityScore();
    }

    public class LayoutSuggestion
    {
        public int OptionNumber { get; set; }
        public string Title { get; set; } = string.Empty;
        public LayoutTemplatePattern Pattern { get; set; }
        public string Reasoning { get; set; } = string.Empty;
        public int ObjectCount { get; set; }
        public string ContentDensity { get; set; } = "Balanced";
        public string DesignImpact { get; set; } = "High Clarity";
        public string IconSymbol { get; set; } = "📊";
        public LayoutQualityScore QualityScore { get; set; } = new LayoutQualityScore();
        public LayoutSpecification Specification { get; set; } = new LayoutSpecification();
    }
}

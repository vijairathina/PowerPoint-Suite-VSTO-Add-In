using System;
using System.Runtime.Serialization;
using OneAssist.SlideCompare.Models;

namespace OneAssist.SlideCompare.Services
{
    /// <summary>
    /// Configurable settings for the slide comparison engine.
    /// </summary>
    [DataContract]
    public class CompareSettings
    {
        [DataMember]
        public bool CompareText { get; set; } = true;

        [DataMember]
        public bool CompareShapes { get; set; } = true;

        [DataMember]
        public bool CompareTables { get; set; } = true;

        [DataMember]
        public bool CompareCharts { get; set; } = true;

        [DataMember]
        public bool CompareImages { get; set; } = true;

        [DataMember]
        public bool CompareFormatting { get; set; } = true;

        [DataMember]
        public bool ComparePosition { get; set; } = true;

        [DataMember]
        public bool CompareSpeakerNotes { get; set; } = false;

        [DataMember]
        public bool CompareHiddenSlides { get; set; } = false;

        [DataMember]
        public double SimilarityThreshold { get; set; } = 0.70; // 70%

        [DataMember]
        public float PositionTolerancePoints { get; set; } = 10.0f; // 10 pt

        [DataMember]
        public float SizeTolerancePoints { get; set; } = 8.0f;

        [DataMember]
        public TextDiffMode TextMatchingMode { get; set; } = TextDiffMode.WordLevel;

        // Configurable Shape Weights (Default to Req 6: 40%, 15%, 15%, 10%, 10%, 10%)
        [DataMember]
        public double WeightText { get; set; } = 0.40;

        [DataMember]
        public double WeightShapeType { get; set; } = 0.15;

        [DataMember]
        public double WeightPosition { get; set; } = 0.15;

        [DataMember]
        public double WeightSize { get; set; } = 0.10;

        [DataMember]
        public double WeightNamePlaceholder { get; set; } = 0.10;

        [DataMember]
        public double WeightVisualContent { get; set; } = 0.10;

        public CompareSettings Clone()
        {
            return (CompareSettings)this.MemberwiseClone();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace OneAssist.SlideCompare.Models
{
    /// <summary>
    /// Independent POCO snapshot of an entire presentation.
    /// </summary>
    [DataContract]
    public class PresentationSnapshot
    {
        [DataMember]
        public string FilePath { get; set; } = string.Empty;

        [DataMember]
        public string PresentationName { get; set; } = string.Empty;

        [DataMember]
        public float SlideWidth { get; set; }

        [DataMember]
        public float SlideHeight { get; set; }

        [DataMember]
        public DateTime ExtractedAt { get; set; } = DateTime.UtcNow;

        [DataMember]
        public List<SlideSnapshot> Slides { get; set; } = new List<SlideSnapshot>();
    }
}

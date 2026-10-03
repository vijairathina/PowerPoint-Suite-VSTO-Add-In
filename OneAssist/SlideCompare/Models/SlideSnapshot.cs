using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace OneAssist.SlideCompare.Models
{
    /// <summary>
    /// Independent POCO snapshot of a PowerPoint slide.
    /// </summary>
    [DataContract]
    public class SlideSnapshot
    {
        [DataMember]
        public int SlideNumber { get; set; }

        [DataMember]
        public int SlideId { get; set; }

        [DataMember]
        public string Title { get; set; } = string.Empty;

        [DataMember]
        public string NormalizedTitle { get; set; } = string.Empty;

        [DataMember]
        public string BodyText { get; set; } = string.Empty;

        [DataMember]
        public string NormalizedBodyText { get; set; } = string.Empty;

        [DataMember]
        public string SpeakerNotes { get; set; } = string.Empty;

        [DataMember]
        public bool IsHidden { get; set; }

        [DataMember]
        public string Fingerprint { get; set; } = string.Empty;

        [DataMember]
        public List<ShapeSnapshot> Shapes { get; set; } = new List<ShapeSnapshot>();

        /// <summary>
        /// Flattens all shapes including child shapes inside groups for searching / comparison.
        /// </summary>
        public List<ShapeSnapshot> GetAllShapesFlattened()
        {
            var list = new List<ShapeSnapshot>();
            if (Shapes == null) return list;

            void Collect(IEnumerable<ShapeSnapshot> collection)
            {
                foreach (var shp in collection)
                {
                    list.Add(shp);
                    if (shp.IsGroup && shp.ChildShapes != null && shp.ChildShapes.Count > 0)
                    {
                        Collect(shp.ChildShapes);
                    }
                }
            }

            Collect(Shapes);
            return list;
        }
    }
}

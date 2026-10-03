namespace OneAssist.ImageEditor.Core
{
    /// <summary>
    /// Stores the geometry and metadata of the selected PowerPoint shape so it can be restored perfectly on Apply.
    /// </summary>
    public class ShapeMetadata
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public float Rotation { get; set; }
        public int ZOrderPosition { get; set; }
        public string Name { get; set; }
        public string Title { get; set; }
        public string AlternativeText { get; set; }
        public int SlideIndex { get; set; }
        public int ShapeId { get; set; }
        public bool IsValid { get; set; } = true;
    }
}

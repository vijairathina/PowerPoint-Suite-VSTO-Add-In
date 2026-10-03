using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace OneAssist.SlideCompare.Models
{
    /// <summary>
    /// Snapshot of a single table cell.
    /// </summary>
    [DataContract]
    public class TableCellSnapshot
    {
        [DataMember]
        public int RowIndex { get; set; }

        [DataMember]
        public int ColumnIndex { get; set; }

        [DataMember]
        public string Text { get; set; } = string.Empty;

        [DataMember]
        public string NormalizedText { get; set; } = string.Empty;
    }

    /// <summary>
    /// Independent POCO snapshot of a PowerPoint shape.
    /// </summary>
    [DataContract]
    public class ShapeSnapshot
    {
        [DataMember]
        public int Id { get; set; }

        [DataMember]
        public int Index { get; set; }

        [DataMember]
        public string Name { get; set; } = string.Empty;

        [DataMember]
        public string ShapeType { get; set; } = "AutoShape";

        [DataMember]
        public string PlaceholderType { get; set; } = string.Empty;

        [DataMember]
        public bool IsPlaceholder { get; set; }

        // Geometry & Layout
        [DataMember]
        public float Left { get; set; }

        [DataMember]
        public float Top { get; set; }

        [DataMember]
        public float Width { get; set; }

        [DataMember]
        public float Height { get; set; }

        [DataMember]
        public float Rotation { get; set; }

        [DataMember]
        public int ZOrderPosition { get; set; }

        // Content
        [DataMember]
        public bool HasText { get; set; }

        [DataMember]
        public string RawText { get; set; } = string.Empty;

        [DataMember]
        public string NormalizedText { get; set; } = string.Empty;

        [DataMember]
        public string Fingerprint { get; set; } = string.Empty;

        // Formatting
        [DataMember]
        public string FontFamily { get; set; } = string.Empty;

        [DataMember]
        public float FontSize { get; set; }

        [DataMember]
        public bool IsBold { get; set; }

        [DataMember]
        public bool IsItalic { get; set; }

        [DataMember]
        public bool IsUnderline { get; set; }

        [DataMember]
        public string FontColorHex { get; set; } = string.Empty;

        [DataMember]
        public string FillColorHex { get; set; } = string.Empty;

        [DataMember]
        public string LineColorHex { get; set; } = string.Empty;

        [DataMember]
        public string TextAlignment { get; set; } = string.Empty;

        // Group support
        [DataMember]
        public bool IsGroup { get; set; }

        [DataMember]
        public List<ShapeSnapshot> ChildShapes { get; set; } = new List<ShapeSnapshot>();

        // Table support
        [DataMember]
        public bool IsTable { get; set; }

        [DataMember]
        public int TableRowCount { get; set; }

        [DataMember]
        public int TableColumnCount { get; set; }

        [DataMember]
        public List<TableCellSnapshot> TableCells { get; set; } = new List<TableCellSnapshot>();

        // Chart support
        [DataMember]
        public bool IsChart { get; set; }

        [DataMember]
        public string ChartTitle { get; set; } = string.Empty;

        [DataMember]
        public List<string> ChartSeriesNames { get; set; } = new List<string>();

        [DataMember]
        public List<string> ChartCategories { get; set; } = new List<string>();

        [DataMember]
        public List<string> ChartValues { get; set; } = new List<string>();

        // Image / Media support
        [DataMember]
        public bool IsImage { get; set; }

        [DataMember]
        public string ImageHash { get; set; } = string.Empty;

        [DataMember]
        public float ImageAspectWidth { get; set; }

        [DataMember]
        public float ImageAspectHeight { get; set; }

        // SmartArt
        [DataMember]
        public bool HasSmartArt { get; set; }

        [DataMember]
        public string SmartArtText { get; set; } = string.Empty;
    }
}

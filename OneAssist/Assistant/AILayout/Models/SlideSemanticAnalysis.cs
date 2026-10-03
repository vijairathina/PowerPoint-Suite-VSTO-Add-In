using System;
using System.Collections.Generic;
using System.Linq;

namespace OneAssist.Assistant.AILayout.Models
{
    public class SemanticShape
    {
        public int ShapeId { get; set; }
        public string ShapeName { get; set; }
        public ObjectSemanticType SemanticType { get; set; } = ObjectSemanticType.Unknown;
        public string RawText { get; set; } = string.Empty;
        public int WordCount { get; set; }
        public int BulletCount { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool IsBold { get; set; }
        public int FillColorBgr { get; set; }
        public int FontColorBgr { get; set; }
        public bool HasChart { get; set; }
        public bool HasTable { get; set; }
        public bool HasPicture { get; set; }
        public int ZOrderPosition { get; set; }
    }

    public class SlideSemanticAnalysis
    {
        public int SlideIndex { get; set; }
        public string SlideTitle { get; set; } = string.Empty;
        public SlideContentType DetectedContentType { get; set; } = SlideContentType.GeneralContent;
        public string ContentTypeReason { get; set; } = string.Empty;
        public List<SemanticShape> SemanticShapes { get; set; } = new List<SemanticShape>();

        public int TotalWordCount => SemanticShapes.Sum(s => s.WordCount);
        public int TotalBulletCount => SemanticShapes.Sum(s => s.BulletCount);
        public int TotalImagesCount => SemanticShapes.Count(s => s.HasPicture || s.SemanticType == ObjectSemanticType.HeroImage || s.SemanticType == ObjectSemanticType.SupportingImage || s.SemanticType == ObjectSemanticType.Logo || s.SemanticType == ObjectSemanticType.Icon);
        public int TotalChartsCount => SemanticShapes.Count(s => s.HasChart || s.SemanticType == ObjectSemanticType.PrimaryChart || s.SemanticType == ObjectSemanticType.SupportingChart);
        public int TotalTablesCount => SemanticShapes.Count(s => s.HasTable || s.SemanticType == ObjectSemanticType.Table);
        public int TotalShapesCount => SemanticShapes.Count;

        public float SlideWidth { get; set; } = 960f; // 13.33 in in points (16:9)
        public float SlideHeight { get; set; } = 540f; // 7.5 in in points

        public List<string> DetectedIssues { get; set; } = new List<string>();
        public List<string> DetectedStructureSummary { get; set; } = new List<string>();

        public SemanticShape TitleShape => SemanticShapes.FirstOrDefault(s => s.SemanticType == ObjectSemanticType.Title);
        public List<SemanticShape> BodyShapes => SemanticShapes.Where(s => s.SemanticType == ObjectSemanticType.Body || s.SemanticType == ObjectSemanticType.Description).ToList();
        public List<SemanticShape> KPIShapes => SemanticShapes.Where(s => s.SemanticType == ObjectSemanticType.KPI).ToList();
        public List<SemanticShape> ImageShapes => SemanticShapes.Where(s => s.HasPicture || s.SemanticType == ObjectSemanticType.HeroImage || s.SemanticType == ObjectSemanticType.SupportingImage).ToList();
        public List<SemanticShape> ChartShapes => SemanticShapes.Where(s => s.HasChart || s.SemanticType == ObjectSemanticType.PrimaryChart || s.SemanticType == ObjectSemanticType.SupportingChart).ToList();
        public List<SemanticShape> TableShapes => SemanticShapes.Where(s => s.HasTable || s.SemanticType == ObjectSemanticType.Table).ToList();
    }
}

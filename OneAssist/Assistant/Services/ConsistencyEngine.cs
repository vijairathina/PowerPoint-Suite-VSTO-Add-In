using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public class SlideTitleMetadata
    {
        public int SlideIndex { get; set; }
        public int ShapeId { get; set; }
        public string ShapeName { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
        public string FontName { get; set; }
        public float? FontSize { get; set; }
    }

    public static class ConsistencyEngine
    {
        public static SlideTitleMetadata DetectSlideTitle(PowerPoint.Slide slide, int slideIndex)
        {
            if (slide == null) return null;
            try
            {
                // First check standard PowerPoint title shape
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    try
                    {
                        if (s.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                        {
                            if (s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderTitle ||
                                s.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle)
                            {
                                return CreateMetadata(s, slideIndex);
                            }
                        }
                    }
                    catch { }
                }

                // Fallback: uppermost text box with substantial font size
                PowerPoint.Shape bestCandidate = null;
                float highestTop = float.MaxValue;
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (SafeInteropHelper.HasText(s) && s.Top < highestTop && s.Top < 200)
                    {
                        float? sz = SafeInteropHelper.GetFontSize(s);
                        if (sz.HasValue && sz.Value >= 18)
                        {
                            highestTop = s.Top;
                            bestCandidate = s;
                        }
                    }
                }

                if (bestCandidate != null)
                {
                    return CreateMetadata(bestCandidate, slideIndex);
                }
            }
            catch { }

            return null;
        }

        private static SlideTitleMetadata CreateMetadata(PowerPoint.Shape s, int slideIndex)
        {
            return new SlideTitleMetadata
            {
                SlideIndex = slideIndex,
                ShapeId = s.Id,
                ShapeName = s.Name,
                Left = s.Left,
                Top = s.Top,
                Width = s.Width,
                Height = s.Height,
                FontName = SafeInteropHelper.GetFontName(s),
                FontSize = SafeInteropHelper.GetFontSize(s)
            };
        }

        public static void SyncTitleWithReference(PowerPoint.Presentation pres, SlideTitleMetadata referenceTitle, int targetSlideIndex)
        {
            if (pres == null || referenceTitle == null) return;
            try
            {
                var targetSlide = pres.Slides[targetSlideIndex];
                var targetMeta = DetectSlideTitle(targetSlide, targetSlideIndex);
                if (targetMeta != null)
                {
                    foreach (PowerPoint.Shape s in targetSlide.Shapes)
                    {
                        if (s.Id == targetMeta.ShapeId)
                        {
                            s.Left = referenceTitle.Left;
                            s.Top = referenceTitle.Top;
                            s.Width = referenceTitle.Width;
                            if (!string.IsNullOrEmpty(referenceTitle.FontName))
                            {
                                s.TextFrame.TextRange.Font.Name = referenceTitle.FontName;
                            }
                            if (referenceTitle.FontSize.HasValue)
                            {
                                s.TextFrame.TextRange.Font.Size = referenceTitle.FontSize.Value;
                            }
                            break;
                        }
                    }
                }
            }
            catch { }
        }
    }
}

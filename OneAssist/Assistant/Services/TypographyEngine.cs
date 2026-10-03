using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Settings;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public class TypographyStats
    {
        public Dictionary<string, int> FontFamilies { get; set; } = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public List<float> FontSizes { get; set; } = new List<float>();
        public int TotalTextShapes { get; set; }
    }

    public static class TypographyEngine
    {
        public static TypographyStats AnalyzeSlideTypography(PowerPoint.Slide slide)
        {
            var stats = new TypographyStats();
            if (slide == null) return stats;

            try
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (SafeInteropHelper.HasText(shape))
                    {
                        stats.TotalTextShapes++;
                        string fontName = SafeInteropHelper.GetFontName(shape);
                        if (!string.IsNullOrEmpty(fontName))
                        {
                            if (!stats.FontFamilies.ContainsKey(fontName))
                                stats.FontFamilies[fontName] = 0;
                            stats.FontFamilies[fontName]++;
                        }

                        float? size = SafeInteropHelper.GetFontSize(shape);
                        if (size.HasValue)
                        {
                            stats.FontSizes.Add(size.Value);
                        }
                    }
                }
            }
            catch { }

            return stats;
        }

        public static void ApplyFont(PowerPoint.Shape shape, string targetFont)
        {
            if (!SafeInteropHelper.HasText(shape) || string.IsNullOrEmpty(targetFont)) return;
            try
            {
                shape.TextFrame.TextRange.Font.Name = targetFont;
            }
            catch { }
        }

        public static void NormalizeMargins(PowerPoint.Shape shape, float marginPoints = 7.2f)
        {
            if (shape == null || shape.HasTextFrame != Office.MsoTriState.msoTrue) return;
            try
            {
                var tf = shape.TextFrame;
                tf.MarginLeft = marginPoints;
                tf.MarginRight = marginPoints;
                tf.MarginTop = marginPoints;
                tf.MarginBottom = marginPoints;
            }
            catch { }
        }
    }
}

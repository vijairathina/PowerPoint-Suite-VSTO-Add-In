using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public class ReplaceStats
    {
        public int MatchesCount { get; set; }
        public int AffectedShapesCount { get; set; }
        public int AffectedSlidesCount { get; set; }
    }

    public static class SmartReplaceService
    {
        /// <summary>
        /// Replaces all occurrences of an old font name with a new font name across the entire presentation.
        /// </summary>
        public static ReplaceStats ReplaceFontGlobally(PowerPoint.Presentation pres, string oldFont, string newFont)
        {
            var stats = new ReplaceStats();
            if (pres == null || string.IsNullOrWhiteSpace(oldFont) || string.IsNullOrWhiteSpace(newFont)) return stats;

            try
            {
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    bool slideModified = false;
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (SafeInteropHelper.HasText(shape))
                        {
                            string curFont = SafeInteropHelper.GetFontName(shape);
                            if (string.Equals(curFont, oldFont, StringComparison.OrdinalIgnoreCase))
                            {
                                shape.TextFrame.TextRange.Font.Name = newFont;
                                stats.AffectedShapesCount++;
                                stats.MatchesCount++;
                                slideModified = true;
                            }
                        }
                    }
                    if (slideModified) stats.AffectedSlidesCount++;
                }
            }
            catch { }

            return stats;
        }

        /// <summary>
        /// Replaces text occurrences across the entire presentation.
        /// </summary>
        public static ReplaceStats ReplaceTextGlobally(PowerPoint.Presentation pres, string findText, string replaceText, bool matchCase)
        {
            var stats = new ReplaceStats();
            if (pres == null || string.IsNullOrEmpty(findText)) return stats;

            try
            {
                var options = matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                string pattern = Regex.Escape(findText);

                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    bool slideModified = false;
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (SafeInteropHelper.HasText(shape))
                        {
                            string text = shape.TextFrame.TextRange.Text;
                            if (!string.IsNullOrEmpty(text) && Regex.IsMatch(text, pattern, options))
                            {
                                int matches = Regex.Matches(text, pattern, options).Count;
                                string replaced = Regex.Replace(text, pattern, replaceText ?? string.Empty, options);
                                shape.TextFrame.TextRange.Text = replaced;

                                stats.MatchesCount += matches;
                                stats.AffectedShapesCount++;
                                slideModified = true;
                            }
                        }
                    }
                    if (slideModified) stats.AffectedSlidesCount++;
                }
            }
            catch { }

            return stats;
        }

        /// <summary>
        /// Replaces colors across the entire presentation.
        /// </summary>
        public static ReplaceStats ReplaceColorGlobally(PowerPoint.Presentation pres, Color oldColor, Color newColor)
        {
            var stats = new ReplaceStats();
            if (pres == null) return stats;

            try
            {
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    int count = ColorEngine.ReplaceColorInSlide(slide, oldColor, newColor);
                    if (count > 0)
                    {
                        stats.MatchesCount += count;
                        stats.AffectedSlidesCount++;
                    }
                }
            }
            catch { }

            return stats;
        }
    }
}

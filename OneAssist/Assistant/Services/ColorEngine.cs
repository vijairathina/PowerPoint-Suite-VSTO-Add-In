using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using OneAssist.Assistant.Settings;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public class ColorUsage
    {
        public string Hex { get; set; }
        public Color Color { get; set; }
        public int UsageCount { get; set; }
        public List<string> Locations { get; set; } = new List<string>();

        public string DisplayText => $"{Hex} ({UsageCount}x)";
    }

    public static class ColorEngine
    {
        /// <summary>
        /// Calculates the relative luminance of a color according to WCAG 2.1 specifications.
        /// </summary>
        public static double CalculateLuminance(Color color)
        {
            double r = ConvertChannel(color.R / 255.0);
            double g = ConvertChannel(color.G / 255.0);
            double b = ConvertChannel(color.B / 255.0);
            return 0.2126 * r + 0.7152 * g + 0.0722 * b;
        }

        private static double ConvertChannel(double channel)
        {
            return channel <= 0.03928 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);
        }

        /// <summary>
        /// Calculates contrast ratio between two colors (ranges from 1.0 to 21.0).
        /// </summary>
        public static double CalculateContrastRatio(Color color1, Color color2)
        {
            double l1 = CalculateLuminance(color1);
            double l2 = CalculateLuminance(color2);
            double brightest = Math.Max(l1, l2);
            double darkest = Math.Min(l1, l2);
            return Math.Round((brightest + 0.05) / (darkest + 0.05), 2);
        }

        /// <summary>
        /// Scans all shapes in a slide and accumulates color usages for Fills, Borders, and Text.
        /// </summary>
        public static Dictionary<string, ColorUsage> ExtractSlideColors(PowerPoint.Slide slide, int slideIndex)
        {
            var result = new Dictionary<string, ColorUsage>(StringComparer.OrdinalIgnoreCase);
            if (slide == null) return result;

            try
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        // Check Fill
                        Color? fill = SafeInteropHelper.GetFillColor(shape);
                        if (fill.HasValue)
                        {
                            RegisterColor(result, fill.Value, $"Slide {slideIndex} - {shape.Name} (Fill)");
                        }

                        // Check Line
                        Color? line = SafeInteropHelper.GetLineColor(shape);
                        if (line.HasValue)
                        {
                            RegisterColor(result, line.Value, $"Slide {slideIndex} - {shape.Name} (Border)");
                        }

                        // Check Text
                        Color? text = SafeInteropHelper.GetTextColor(shape);
                        if (text.HasValue)
                        {
                            RegisterColor(result, text.Value, $"Slide {slideIndex} - {shape.Name} (Text)");
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return result;
        }

        private static void RegisterColor(Dictionary<string, ColorUsage> dict, Color color, string location)
        {
            string hex = SafeInteropHelper.ColorToHex(color);
            if (!dict.TryGetValue(hex, out var usage))
            {
                usage = new ColorUsage
                {
                    Hex = hex,
                    Color = color,
                    UsageCount = 0
                };
                dict[hex] = usage;
            }
            usage.UsageCount++;
            if (usage.Locations.Count < 50)
            {
                usage.Locations.Add(location);
            }
        }

        /// <summary>
        /// Replaces an old color with a new color across shapes in a slide or presentation.
        /// </summary>
        public static int ReplaceColorInSlide(PowerPoint.Slide slide, Color oldColor, Color newColor)
        {
            if (slide == null) return 0;
            int replacedCount = 0;
            int oldBgr = SafeInteropHelper.ColorToPowerPointRgb(oldColor);
            int newBgr = SafeInteropHelper.ColorToPowerPointRgb(newColor);

            try
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        // Replace Fill
                        if (shape.Fill != null && shape.Fill.Visible == Office.MsoTriState.msoTrue)
                        {
                            if (shape.Fill.ForeColor.RGB == oldBgr)
                            {
                                shape.Fill.ForeColor.RGB = newBgr;
                                replacedCount++;
                            }
                        }

                        // Replace Line
                        if (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue)
                        {
                            if (shape.Line.ForeColor.RGB == oldBgr)
                            {
                                shape.Line.ForeColor.RGB = newBgr;
                                replacedCount++;
                            }
                        }

                        // Replace Text
                        if (SafeInteropHelper.HasText(shape))
                        {
                            var tr = shape.TextFrame.TextRange;
                            if (tr.Font.Color.RGB == oldBgr)
                            {
                                tr.Font.Color.RGB = newBgr;
                                replacedCount++;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return replacedCount;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OneAssist.Assistant.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Tokens
{
    public class DesignTokensModel
    {
        public string TitleFont { get; set; } = "Aptos Display";
        public float TitleFontSize { get; set; } = 28.0f;
        public string TitleColorHex { get; set; } = "#002060";

        public string BodyFont { get; set; } = "Aptos";
        public float BodyFontSize { get; set; } = 14.0f;
        public string BodyColorHex { get; set; } = "#262626";

        public string PrimaryColorHex { get; set; } = "#002060";
        public string SecondaryColorHex { get; set; } = "#00A88F";
        public string AccentColorHex { get; set; } = "#FFB900";

        public float SpacingSmall { get; set; } = 7.2f;   // 0.1 inch
        public float SpacingMedium { get; set; } = 14.4f; // 0.2 inch
        public float SpacingLarge { get; set; } = 28.8f;  // 0.4 inch
    }

    public class DesignSystemReport
    {
        public List<string> DetectedFonts { get; set; } = new List<string>();
        public List<string> DetectedColors { get; set; } = new List<string>();
        public List<float> DetectedFontSizes { get; set; } = new List<float>();
        public int TotalSlidesScanned { get; set; }
        public int TotalShapesScanned { get; set; }

        public string GenerateMarkdownSummary()
        {
            var sb = new StringBuilder();
            sb.AppendLine("# Design System Extracted from Presentation");
            sb.AppendLine($"*Scanned {TotalSlidesScanned} slides, {TotalShapesScanned} shapes.*");
            sb.AppendLine();
            sb.AppendLine("### Typography Tokens");
            foreach (var f in DetectedFonts) sb.AppendLine($"- Font Family: `{f}`");
            sb.AppendLine();
            sb.AppendLine("### Color Tokens");
            foreach (var c in DetectedColors) sb.AppendLine($"- Color: `{c}`");
            return sb.ToString();
        }
    }

    public static class DesignTokenService
    {
        public static DesignSystemReport ExtractDesignSystem(PowerPoint.Presentation pres)
        {
            var report = new DesignSystemReport();
            if (pres == null) return report;

            var fontsSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var colorsSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var sizesSet = new HashSet<float>();

            try
            {
                report.TotalSlidesScanned = pres.Slides.Count;
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    foreach (PowerPoint.Shape s in slide.Shapes)
                    {
                        report.TotalShapesScanned++;

                        // Fonts & Sizes
                        if (SafeInteropHelper.HasText(s))
                        {
                            string fn = SafeInteropHelper.GetFontName(s);
                            if (!string.IsNullOrEmpty(fn)) fontsSet.Add(fn);

                            float? sz = SafeInteropHelper.GetFontSize(s);
                            if (sz.HasValue) sizesSet.Add(sz.Value);

                            var tc = SafeInteropHelper.GetTextColor(s);
                            if (tc.HasValue) colorsSet.Add(SafeInteropHelper.ColorToHex(tc.Value));
                        }

                        // Fills
                        var fill = SafeInteropHelper.GetFillColor(s);
                        if (fill.HasValue) colorsSet.Add(SafeInteropHelper.ColorToHex(fill.Value));

                        // Lines
                        var line = SafeInteropHelper.GetLineColor(s);
                        if (line.HasValue) colorsSet.Add(SafeInteropHelper.ColorToHex(line.Value));
                    }
                }

                report.DetectedFonts = fontsSet.OrderBy(x => x).ToList();
                report.DetectedColors = colorsSet.OrderBy(x => x).ToList();
                report.DetectedFontSizes = sizesSet.OrderByDescending(x => x).ToList();
            }
            catch { }

            return report;
        }

        public static void ApplyTokensToSlide(PowerPoint.Slide slide, DesignTokensModel tokens)
        {
            if (slide == null || tokens == null) return;
            try
            {
                var titleColor = SafeInteropHelper.HexToColor(tokens.TitleColorHex);
                var bodyColor = SafeInteropHelper.HexToColor(tokens.BodyColorHex);

                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (SafeInteropHelper.HasText(s))
                    {
                        var tf = s.TextFrame.TextRange;
                        float? sz = SafeInteropHelper.GetFontSize(s);

                        if (sz.HasValue && sz.Value >= 20.0f)
                        {
                            // Title
                            tf.Font.Name = tokens.TitleFont;
                            tf.Font.Size = tokens.TitleFontSize;
                            tf.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(titleColor);
                        }
                        else
                        {
                            // Body
                            tf.Font.Name = tokens.BodyFont;
                            tf.Font.Size = tokens.BodyFontSize;
                            tf.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(bodyColor);
                        }
                    }
                }
            }
            catch { }
        }
    }
}

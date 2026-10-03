using System;
using System.Drawing;
using OneAssist.Assistant.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Grid
{
    public class FooterOptions
    {
        public string ConfidentialityNotice { get; set; } = "Confidential - For Internal Use Only";
        public string CompanyName { get; set; } = "OneAssist Enterprise";
        public bool IncludeSlideNumbers { get; set; } = true;
        public string SourceCitation { get; set; }
    }

    public static class FooterManagerService
    {
        public static void ApplyFooterToSlide(PowerPoint.Slide slide, int currentSlide, int totalSlides, FooterOptions options)
        {
            if (slide == null || options == null) return;

            try
            {
                float sw = slide.Application.ActivePresentation.PageSetup.SlideWidth;
                float sh = slide.Application.ActivePresentation.PageSetup.SlideHeight;
                float footerY = sh - 28.0f;

                // 1. Confidentiality / Company (Bottom Left)
                if (!string.IsNullOrEmpty(options.ConfidentialityNotice))
                {
                    string footerText = $"{options.CompanyName} | {options.ConfidentialityNotice}";
                    var ft = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 36.0f, footerY, 400.0f, 20.0f);
                    ft.Name = "Footer_Confidentiality";
                    ft.TextFrame.TextRange.Text = footerText;
                    ft.TextFrame.TextRange.Font.Name = "Aptos";
                    ft.TextFrame.TextRange.Font.Size = 9.0f;
                    ft.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(148, 163, 184));
                    ft.TextFrame.MarginLeft = 0;
                    ft.TextFrame.MarginBottom = 0;
                }

                // 2. Slide Number (Bottom Right: "03 / 24")
                if (options.IncludeSlideNumbers && totalSlides > 0)
                {
                    string numText = $"{currentSlide:D2} / {totalSlides:D2}";
                    var sn = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, sw - 120.0f, footerY, 84.0f, 20.0f);
                    sn.Name = "Footer_SlideNumber";
                    sn.TextFrame.TextRange.Text = numText;
                    sn.TextFrame.TextRange.Font.Name = "Aptos";
                    sn.TextFrame.TextRange.Font.Size = 9.0f;
                    sn.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    sn.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
                    sn.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(148, 163, 184));
                }

                // 3. Source Footnote (Right above footer if specified)
                if (!string.IsNullOrEmpty(options.SourceCitation))
                {
                    var src = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 36.0f, footerY - 18.0f, sw - 72.0f, 18.0f);
                    src.Name = "Footnote_Source";
                    src.TextFrame.TextRange.Text = $"Source: {options.SourceCitation}";
                    src.TextFrame.TextRange.Font.Name = "Aptos";
                    src.TextFrame.TextRange.Font.Size = 8.5f;
                    src.TextFrame.TextRange.Font.Italic = Office.MsoTriState.msoTrue;
                    src.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(100, 116, 139));
                }
            }
            catch { }
        }
    }
}

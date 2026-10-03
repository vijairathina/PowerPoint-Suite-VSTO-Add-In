using System;
using System.Drawing;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public class FormatPainterOptions
    {
        public bool CopyFontName { get; set; } = true;
        public bool CopyFontSize { get; set; } = true;
        public bool CopyFontColor { get; set; } = true;
        public bool CopyFontStyles { get; set; } = true; // Bold, Italic
        public bool CopyFill { get; set; } = true;
        public bool CopyBorder { get; set; } = true;
        public bool CopyMargins { get; set; } = true;
        public bool CopySize { get; set; } = false;
    }

    public class FormatSnapshot
    {
        public string FontName { get; set; }
        public float? FontSize { get; set; }
        public Color? FontColor { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }

        public bool HasFill { get; set; }
        public Color? FillColor { get; set; }

        public bool HasLine { get; set; }
        public Color? LineColor { get; set; }
        public float LineWeight { get; set; }

        public float MarginLeft { get; set; } = 7.2f;
        public float MarginTop { get; set; } = 7.2f;
        public float MarginRight { get; set; } = 7.2f;
        public float MarginBottom { get; set; } = 7.2f;

        public float Width { get; set; }
        public float Height { get; set; }
    }

    public static class FormatPainterService
    {
        public static FormatSnapshot CaptureFormatting(PowerPoint.Shape source)
        {
            if (source == null) return null;
            var snap = new FormatSnapshot
            {
                Width = source.Width,
                Height = source.Height
            };

            try
            {
                // Font formatting
                if (SafeInteropHelper.HasText(source))
                {
                    var font = source.TextFrame.TextRange.Font;
                    snap.FontName = font.Name;
                    snap.FontSize = font.Size;
                    snap.FontColor = SafeInteropHelper.GetTextColor(source);
                    snap.IsBold = font.Bold == Office.MsoTriState.msoTrue;
                    snap.IsItalic = font.Italic == Office.MsoTriState.msoTrue;

                    var tf = source.TextFrame;
                    snap.MarginLeft = tf.MarginLeft;
                    snap.MarginTop = tf.MarginTop;
                    snap.MarginRight = tf.MarginRight;
                    snap.MarginBottom = tf.MarginBottom;
                }

                // Fill formatting
                if (source.Fill != null && source.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    snap.HasFill = true;
                    snap.FillColor = SafeInteropHelper.GetFillColor(source);
                }

                // Border/Line formatting
                if (source.Line != null && source.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    snap.HasLine = true;
                    snap.LineColor = SafeInteropHelper.GetLineColor(source);
                    snap.LineWeight = source.Line.Weight;
                }
            }
            catch { }

            return snap;
        }

        public static void ApplyFormatting(PowerPoint.Shape target, FormatSnapshot snapshot, FormatPainterOptions options)
        {
            if (target == null || snapshot == null || options == null) return;

            try
            {
                // Apply Text Formatting
                if (SafeInteropHelper.HasText(target))
                {
                    var font = target.TextFrame.TextRange.Font;
                    if (options.CopyFontName && !string.IsNullOrEmpty(snapshot.FontName))
                    {
                        font.Name = snapshot.FontName;
                    }
                    if (options.CopyFontSize && snapshot.FontSize.HasValue)
                    {
                        font.Size = snapshot.FontSize.Value;
                    }
                    if (options.CopyFontColor && snapshot.FontColor.HasValue)
                    {
                        font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(snapshot.FontColor.Value);
                    }
                    if (options.CopyFontStyles)
                    {
                        font.Bold = snapshot.IsBold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                        font.Italic = snapshot.IsItalic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    }

                    if (options.CopyMargins)
                    {
                        var tf = target.TextFrame;
                        tf.MarginLeft = snapshot.MarginLeft;
                        tf.MarginTop = snapshot.MarginTop;
                        tf.MarginRight = snapshot.MarginRight;
                        tf.MarginBottom = snapshot.MarginBottom;
                    }
                }

                // Apply Fill Formatting
                if (options.CopyFill && snapshot.HasFill && snapshot.FillColor.HasValue)
                {
                    target.Fill.Visible = Office.MsoTriState.msoTrue;
                    target.Fill.Solid();
                    target.Fill.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(snapshot.FillColor.Value);
                }

                // Apply Border Formatting
                if (options.CopyBorder && snapshot.HasLine && snapshot.LineColor.HasValue)
                {
                    target.Line.Visible = Office.MsoTriState.msoTrue;
                    target.Line.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(snapshot.LineColor.Value);
                    if (snapshot.LineWeight > 0)
                    {
                        target.Line.Weight = snapshot.LineWeight;
                    }
                }

                // Apply Size
                if (options.CopySize && snapshot.Width > 0 && snapshot.Height > 0)
                {
                    target.Width = snapshot.Width;
                    target.Height = snapshot.Height;
                }
            }
            catch { }
        }
    }
}

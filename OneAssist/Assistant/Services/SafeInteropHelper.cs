using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    /// <summary>
    /// Defensive helper for safely reading and writing PowerPoint COM Interop properties.
    /// Handles BGR/RGB color conversions, COM exceptions, and memory cleanup.
    /// </summary>
    public static class SafeInteropHelper
    {
        /// <summary>
        /// Converts a Win32 BGR COLORREF integer from PowerPoint to a System.Drawing.Color.
        /// </summary>
        public static Color ColorFromPowerPointRgb(int bgr)
        {
            int r = bgr & 0xFF;
            int g = (bgr >> 8) & 0xFF;
            int b = (bgr >> 16) & 0xFF;
            return Color.FromArgb(r, g, b);
        }

        /// <summary>
        /// Converts a System.Drawing.Color to a Win32 BGR COLORREF integer expected by PowerPoint.
        /// </summary>
        public static int ColorToPowerPointRgb(Color color)
        {
            return color.R | (color.G << 8) | (color.B << 16);
        }

        public static string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        public static Color HexToColor(string hex)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(hex)) return Color.Black;
                if (!hex.StartsWith("#")) hex = "#" + hex;
                return ColorTranslator.FromHtml(hex);
            }
            catch
            {
                return Color.Black;
            }
        }

        /// <summary>
        /// Safely checks if a shape contains text.
        /// </summary>
        public static bool HasText(PowerPoint.Shape shape)
        {
            if (shape == null) return false;
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame != null)
                {
                    return shape.TextFrame.HasText == Office.MsoTriState.msoTrue;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Safely retrieves text content from a shape.
        /// </summary>
        public static string GetText(PowerPoint.Shape shape)
        {
            if (!HasText(shape)) return string.Empty;
            try
            {
                return shape.TextFrame.TextRange.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Safely retrieves the primary font name of a shape.
        /// </summary>
        public static string GetFontName(PowerPoint.Shape shape)
        {
            if (!HasText(shape)) return null;
            try
            {
                return shape.TextFrame.TextRange.Font.Name;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Safely retrieves the font size in points.
        /// </summary>
        public static float? GetFontSize(PowerPoint.Shape shape)
        {
            if (!HasText(shape)) return null;
            try
            {
                return shape.TextFrame.TextRange.Font.Size;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Safely retrieves text color.
        /// </summary>
        public static Color? GetTextColor(PowerPoint.Shape shape)
        {
            if (!HasText(shape)) return null;
            try
            {
                int bgr = shape.TextFrame.TextRange.Font.Color.RGB;
                return ColorFromPowerPointRgb(bgr);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Safely retrieves fill color if shape has a solid fill.
        /// </summary>
        public static Color? GetFillColor(PowerPoint.Shape shape)
        {
            if (shape == null) return null;
            try
            {
                if (shape.Fill != null && shape.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                    {
                        return ColorFromPowerPointRgb(shape.Fill.ForeColor.RGB);
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Safely retrieves line/border color.
        /// </summary>
        public static Color? GetLineColor(PowerPoint.Shape shape)
        {
            if (shape == null) return null;
            try
            {
                if (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    return ColorFromPowerPointRgb(shape.Line.ForeColor.RGB);
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Safely releases a COM object reference.
        /// </summary>
        public static void Release(object comObj)
        {
            if (comObj != null && Marshal.IsComObject(comObj))
            {
                try
                {
                    Marshal.ReleaseComObject(comObj);
                }
                catch { }
            }
        }
    }
}

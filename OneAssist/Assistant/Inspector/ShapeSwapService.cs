using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Inspector
{
    public class ShapeSwapOptions
    {
        public bool SwapPosition { get; set; } = true;
        public bool SwapSize { get; set; } = true;
        public bool SwapFormatting { get; set; } = false;
        public bool SwapText { get; set; } = false;
    }

    public static class ShapeSwapService
    {
        public static bool Swap(PowerPoint.Shape s1, PowerPoint.Shape s2, ShapeSwapOptions options)
        {
            if (s1 == null || s2 == null || options == null) return false;

            try
            {
                // 1. Position
                if (options.SwapPosition)
                {
                    float tempLeft = s1.Left;
                    float tempTop = s1.Top;
                    s1.Left = s2.Left;
                    s1.Top = s2.Top;
                    s2.Left = tempLeft;
                    s2.Top = tempTop;
                }

                // 2. Size
                if (options.SwapSize)
                {
                    float tempW = s1.Width;
                    float tempH = s1.Height;
                    s1.Width = s2.Width;
                    s1.Height = s2.Height;
                    s2.Width = tempW;
                    s2.Height = tempH;
                }

                // 3. Formatting
                if (options.SwapFormatting)
                {
                    var f1 = Services.FormatPainterService.CaptureFormatting(s1);
                    var f2 = Services.FormatPainterService.CaptureFormatting(s2);
                    var opt = new Services.FormatPainterOptions { CopySize = false };
                    Services.FormatPainterService.ApplyFormatting(s1, f2, opt);
                    Services.FormatPainterService.ApplyFormatting(s2, f1, opt);
                }

                // 4. Text
                if (options.SwapText && Services.SafeInteropHelper.HasText(s1) && Services.SafeInteropHelper.HasText(s2))
                {
                    string txt1 = s1.TextFrame.TextRange.Text;
                    string txt2 = s2.TextFrame.TextRange.Text;
                    s1.TextFrame.TextRange.Text = txt2;
                    s2.TextFrame.TextRange.Text = txt1;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}

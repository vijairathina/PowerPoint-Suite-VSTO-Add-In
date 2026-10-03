using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using OneAssist.ImageEditor.Core;
using OneAssist.ImageEditor.Utils;

namespace OneAssist.ImageEditor.Processing
{
    /// <summary>
    /// Renders background layers (Solid, Gradient, Image) and composites with foreground.
    /// </summary>
    public static class BackgroundRenderer
    {
        public static Bitmap RenderComposite(
            Bitmap foreground,
            ImageMask alphaMask,
            AdjustmentSettings adjustments,
            BackgroundSettings background,
            bool drawCheckerboardForTransparency = false)
        {
            if (foreground == null) return null;

            int width = foreground.Width;
            int height = foreground.Height;

            // 1. Prepare adjusted foreground with alpha mask
            var adjustedFg = ImageAdjuster.ApplyAdjustments(foreground, adjustments);
            if (alphaMask != null)
            {
                alphaMask.ApplyToBitmap(adjustedFg);
            }

            // If background is transparent and no checkerboard requested, simply return adjusted foreground
            if (background.Type == BackgroundType.Transparent && !drawCheckerboardForTransparency)
            {
                return adjustedFg;
            }

            // 2. Create composite target
            var composite = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(composite))
            {
                g.CompositingMode = CompositingMode.SourceOver;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                // Render background
                RenderBackgroundLayer(g, width, height, background, drawCheckerboardForTransparency);

                // Render foreground on top
                g.DrawImage(adjustedFg, 0, 0, width, height);
            }

            adjustedFg.Dispose();
            return composite;
        }

        private static void RenderBackgroundLayer(
            Graphics g,
            int width,
            int height,
            BackgroundSettings bg,
            bool drawCheckerboard)
        {
            var rect = new Rectangle(0, 0, width, height);

            switch (bg.Type)
            {
                case BackgroundType.Transparent:
                    if (drawCheckerboard)
                    {
                        DrawCheckerboard(g, rect);
                    }
                    break;

                case BackgroundType.White:
                    using (var b = new SolidBrush(Color.White))
                        g.FillRectangle(b, rect);
                    break;

                case BackgroundType.Black:
                    using (var b = new SolidBrush(Color.Black))
                        g.FillRectangle(b, rect);
                    break;

                case BackgroundType.LightGray:
                    using (var b = new SolidBrush(Color.FromArgb(240, 240, 240)))
                        g.FillRectangle(b, rect);
                    break;

                case BackgroundType.CustomColor:
                    using (var b = new SolidBrush(bg.CustomColor))
                        g.FillRectangle(b, rect);
                    break;

                case BackgroundType.LinearGradient:
                    using (var brush = new LinearGradientBrush(rect, bg.GradientColor1, bg.GradientColor2, bg.GradientAngle))
                    {
                        g.FillRectangle(brush, rect);
                    }
                    break;

                case BackgroundType.RadialGradient:
                    using (var path = new GraphicsPath())
                    {
                        path.AddEllipse(rect);
                        using (var pgb = new PathGradientBrush(path))
                        {
                            pgb.CenterColor = bg.GradientColor1;
                            pgb.SurroundColors = new[] { bg.GradientColor2 };
                            g.FillRectangle(pgb, rect);
                        }
                    }
                    break;

                case BackgroundType.Image:
                    if (bg.BackgroundImage != null)
                    {
                        float scale = Math.Max(0.1f, bg.ImageScale);
                        float drawW = bg.BackgroundImage.Width * scale;
                        float drawH = bg.BackgroundImage.Height * scale;
                        float offsetX = (width - drawW) / 2.0f + (bg.ImageOffsetX / 100.0f) * width;
                        float offsetY = (height - drawH) / 2.0f + (bg.ImageOffsetY / 100.0f) * height;

                        g.DrawImage(bg.BackgroundImage, offsetX, offsetY, drawW, drawH);
                    }
                    break;
            }
        }

        private static void DrawCheckerboard(Graphics g, Rectangle rect)
        {
            int cellSize = 16;
            using (var lightBrush = new SolidBrush(Color.FromArgb(250, 250, 250)))
            using (var darkBrush = new SolidBrush(Color.FromArgb(220, 220, 220)))
            {
                for (int y = 0; y < rect.Height; y += cellSize)
                {
                    for (int x = 0; x < rect.Width; x += cellSize)
                    {
                        bool isLight = ((x / cellSize) + (y / cellSize)) % 2 == 0;
                        int w = Math.Min(cellSize, rect.Width - x);
                        int h = Math.Min(cellSize, rect.Height - y);
                        g.FillRectangle(isLight ? lightBrush : darkBrush, x, y, w, h);
                    }
                }
            }
        }
    }
}

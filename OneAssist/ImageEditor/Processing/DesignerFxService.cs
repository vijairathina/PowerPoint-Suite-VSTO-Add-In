using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using OneAssist.ImageEditor.Utils;

namespace OneAssist.ImageEditor.Processing
{
    /// <summary>
    /// Designer FX effects: drop shadow, glow, color tint overlay, vignette, and color temperature.
    /// These are applied on-export (composited to a final bitmap) or as live preview overlays.
    /// </summary>
    public static class DesignerFxService
    {
        // ─────────────────────────────────────────────────────────────────────────────
        // Drop Shadow – renders the foreground with a configurable shadow underneath.
        // ─────────────────────────────────────────────────────────────────────────────
        public static Bitmap ApplyDropShadow(
            Bitmap foreground,
            int offsetX = 8,
            int offsetY = 8,
            int blurRadius = 12,
            float shadowOpacity = 0.55f,
            Color? shadowColor = null)
        {
            if (foreground == null) return null;

            Color sc = shadowColor ?? Color.Black;
            int w = foreground.Width  + Math.Abs(offsetX) + blurRadius * 2;
            int h = foreground.Height + Math.Abs(offsetY) + blurRadius * 2;

            int originX = blurRadius + Math.Max(0, -offsetX);
            int originY = blurRadius + Math.Max(0, -offsetY);
            int shadowX = originX + offsetX;
            int shadowY = originY + offsetY;

            var result = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(result))
            {
                g.Clear(Color.Transparent);

                // Draw shadow layer (grayscale silhouette of the foreground)
                using (var shadowBmp = BuildSilhouette(foreground, sc, shadowOpacity))
                using (var blurredShadow = BoxBlur(shadowBmp, blurRadius))
                {
                    g.DrawImage(blurredShadow, shadowX, shadowY);
                }

                // Draw foreground on top
                g.DrawImage(foreground, originX, originY);
            }

            return result;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Glow Effect – blurred bright halo around the subject
        // ─────────────────────────────────────────────────────────────────────────────
        public static Bitmap ApplyGlow(
            Bitmap foreground,
            int radius = 12,
            float intensity = 0.6f,
            Color? glowColor = null)
        {
            if (foreground == null) return null;

            Color gc = glowColor ?? Color.White;
            int pad = radius * 2;

            var result = new Bitmap(foreground.Width + pad * 2, foreground.Height + pad * 2, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(result))
            {
                g.Clear(Color.Transparent);

                using (var glow = BuildSilhouette(foreground, gc, intensity))
                using (var blurred = BoxBlur(glow, radius))
                {
                    g.DrawImage(blurred, 0, 0);
                }

                g.DrawImage(foreground, pad, pad);
            }

            return result;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Color Tint – blends a color overlay on top of the image
        // ─────────────────────────────────────────────────────────────────────────────
        public static Bitmap ApplyColorTint(Bitmap source, Color tintColor, float strength)
        {
            if (source == null) return null;
            strength = Math.Max(0f, Math.Min(1f, strength));

            var result = BitmapHelper.Clone32bppArgb(source);
            var rect   = new Rectangle(0, 0, result.Width, result.Height);
            var data   = result.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

            unsafe
            {
                byte* ptr = (byte*)data.Scan0;
                int stride = data.Stride;

                for (int y = 0; y < result.Height; y++)
                {
                    byte* row = ptr + y * stride;
                    for (int x = 0; x < result.Width; x++)
                    {
                        int offset = x * 4;
                        float alpha = row[offset + 3] / 255f;
                        if (alpha < 0.01f) continue;

                        row[offset + 0] = ClampByte(row[offset + 0] * (1 - strength) + tintColor.B * strength);
                        row[offset + 1] = ClampByte(row[offset + 1] * (1 - strength) + tintColor.G * strength);
                        row[offset + 2] = ClampByte(row[offset + 2] * (1 - strength) + tintColor.R * strength);
                    }
                }
            }

            result.UnlockBits(data);
            return result;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Vignette – darkens edges toward corners (like a lens/film effect)
        // ─────────────────────────────────────────────────────────────────────────────
        public static Bitmap ApplyVignette(Bitmap source, float strength = 0.5f)
        {
            if (source == null) return null;
            strength = Math.Max(0f, Math.Min(1f, strength));

            var result = BitmapHelper.Clone32bppArgb(source);
            int w = result.Width;
            int h = result.Height;

            float cx = w / 2f;
            float cy = h / 2f;
            float maxR = (float)Math.Sqrt(cx * cx + cy * cy);

            var rect = new Rectangle(0, 0, w, h);
            var data = result.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

            unsafe
            {
                byte* ptr = (byte*)data.Scan0;
                int stride = data.Stride;

                for (int y = 0; y < h; y++)
                {
                    byte* row = ptr + y * stride;
                    for (int x = 0; x < w; x++)
                    {
                        int offset = x * 4;
                        if (row[offset + 3] < 5) continue;

                        float dx = x - cx;
                        float dy = y - cy;
                        float dist = (float)Math.Sqrt(dx * dx + dy * dy) / maxR;

                        // Smooth falloff starting from 40% radius outward
                        float fade = Math.Max(0f, (dist - 0.4f) / 0.6f);
                        float darkening = 1f - fade * fade * strength;

                        row[offset + 0] = ClampByte(row[offset + 0] * darkening);
                        row[offset + 1] = ClampByte(row[offset + 1] * darkening);
                        row[offset + 2] = ClampByte(row[offset + 2] * darkening);
                    }
                }
            }

            result.UnlockBits(data);
            return result;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Color Temperature – warm/cool shift (positive = warmer, negative = cooler)
        // ─────────────────────────────────────────────────────────────────────────────
        public static Bitmap ApplyColorTemperature(Bitmap source, int temperature)
        {
            if (source == null) return null;
            if (temperature == 0) return BitmapHelper.Clone32bppArgb(source);

            float factor = temperature / 100f; // -1.0 to +1.0
            var result = BitmapHelper.Clone32bppArgb(source);
            var rect = new Rectangle(0, 0, result.Width, result.Height);
            var data = result.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

            unsafe
            {
                byte* ptr = (byte*)data.Scan0;
                int stride = data.Stride;

                for (int y = 0; y < result.Height; y++)
                {
                    byte* row = ptr + y * stride;
                    for (int x = 0; x < result.Width; x++)
                    {
                        int offset = x * 4;
                        if (row[offset + 3] < 5) continue;

                        // Warm = more red/yellow, cool = more blue/cyan
                        if (factor > 0)
                        {
                            row[offset + 2] = ClampByte(row[offset + 2] + factor * 30); // R up
                            row[offset + 0] = ClampByte(row[offset + 0] - factor * 15); // B down
                        }
                        else
                        {
                            row[offset + 0] = ClampByte(row[offset + 0] - factor * 30); // B up
                            row[offset + 2] = ClampByte(row[offset + 2] + factor * 15); // R down
                        }
                    }
                }
            }

            result.UnlockBits(data);
            return result;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Helpers
        // ─────────────────────────────────────────────────────────────────────────────
        private static Bitmap BuildSilhouette(Bitmap source, Color color, float opacity)
        {
            var bmp = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                var cm = new ColorMatrix(new float[][]
                {
                    new float[] { 0, 0, 0, 0, color.R / 255f },
                    new float[] { 0, 0, 0, 0, color.G / 255f },
                    new float[] { 0, 0, 0, 0, color.B / 255f },
                    new float[] { 0, 0, 0, opacity, 0 },
                    new float[] { 0, 0, 0, 0, 1 }
                });
                var ia = new ImageAttributes();
                ia.SetColorMatrix(cm);
                g.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height),
                    0, 0, source.Width, source.Height, GraphicsUnit.Pixel, ia);
            }
            return bmp;
        }

        /// <summary>Simple 3-pass box blur approximation of a Gaussian blur.</summary>
        public static Bitmap BoxBlur(Bitmap source, int radius)
        {
            if (radius <= 0) return BitmapHelper.Clone32bppArgb(source);

            var tmp = HorizontalBlur(source, radius);
            var res = VerticalBlur(tmp, radius);
            tmp.Dispose();
            return res;
        }

        private static unsafe Bitmap HorizontalBlur(Bitmap bmp, int r)
        {
            int w = bmp.Width, h = bmp.Height;
            var result = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            var srcData = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var dstData = result.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);

            byte* src = (byte*)srcData.Scan0;
            byte* dst = (byte*)dstData.Scan0;
            int stride = srcData.Stride;

            for (int y = 0; y < h; y++)
            {
                byte* srcRow = src + y * stride;
                byte* dstRow = dst + y * stride;
                int sum0 = 0, sum1 = 0, sum2 = 0, sum3 = 0;
                int count = 0;

                for (int x = -r; x < w; x++)
                {
                    int addIdx = x + r;
                    if (addIdx < w)
                    {
                        sum0 += srcRow[addIdx * 4];
                        sum1 += srcRow[addIdx * 4 + 1];
                        sum2 += srcRow[addIdx * 4 + 2];
                        sum3 += srcRow[addIdx * 4 + 3];
                        count++;
                    }
                    int removeIdx = x - r - 1;
                    if (removeIdx >= 0)
                    {
                        sum0 -= srcRow[removeIdx * 4];
                        sum1 -= srcRow[removeIdx * 4 + 1];
                        sum2 -= srcRow[removeIdx * 4 + 2];
                        sum3 -= srcRow[removeIdx * 4 + 3];
                        count--;
                    }
                    if (x >= 0)
                    {
                        dstRow[x * 4]     = (byte)(sum0 / count);
                        dstRow[x * 4 + 1] = (byte)(sum1 / count);
                        dstRow[x * 4 + 2] = (byte)(sum2 / count);
                        dstRow[x * 4 + 3] = (byte)(sum3 / count);
                    }
                }
            }

            bmp.UnlockBits(srcData);
            result.UnlockBits(dstData);
            return result;
        }

        private static unsafe Bitmap VerticalBlur(Bitmap bmp, int r)
        {
            int w = bmp.Width, h = bmp.Height;
            var result = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            var srcData = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var dstData = result.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);

            byte* src = (byte*)srcData.Scan0;
            byte* dst = (byte*)dstData.Scan0;
            int stride = srcData.Stride;

            for (int x = 0; x < w; x++)
            {
                int sum0 = 0, sum1 = 0, sum2 = 0, sum3 = 0;
                int count = 0;

                for (int y = -r; y < h; y++)
                {
                    int addY = y + r;
                    if (addY < h)
                    {
                        sum0 += src[addY * stride + x * 4];
                        sum1 += src[addY * stride + x * 4 + 1];
                        sum2 += src[addY * stride + x * 4 + 2];
                        sum3 += src[addY * stride + x * 4 + 3];
                        count++;
                    }
                    int removeY = y - r - 1;
                    if (removeY >= 0)
                    {
                        sum0 -= src[removeY * stride + x * 4];
                        sum1 -= src[removeY * stride + x * 4 + 1];
                        sum2 -= src[removeY * stride + x * 4 + 2];
                        sum3 -= src[removeY * stride + x * 4 + 3];
                        count--;
                    }
                    if (y >= 0)
                    {
                        dst[y * stride + x * 4]     = (byte)(sum0 / count);
                        dst[y * stride + x * 4 + 1] = (byte)(sum1 / count);
                        dst[y * stride + x * 4 + 2] = (byte)(sum2 / count);
                        dst[y * stride + x * 4 + 3] = (byte)(sum3 / count);
                    }
                }
            }

            bmp.UnlockBits(srcData);
            result.UnlockBits(dstData);
            return result;
        }

        private static byte ClampByte(float v) => (byte)Math.Max(0, Math.Min(255, (int)v));
    }
}

using System;
using System.Drawing;
using System.Drawing.Imaging;
using OneAssist.ImageEditor.Core;
using OneAssist.ImageEditor.Utils;

namespace OneAssist.ImageEditor.Processing
{
    /// <summary>
    /// Fast pixel-level adjustments: Brightness, Contrast, Saturation, Exposure, Sharpness, Blur, Opacity.
    /// </summary>
    public static class ImageAdjuster
    {
        public static Bitmap ApplyAdjustments(Bitmap source, AdjustmentSettings adj)
        {
            if (source == null) return null;
            if (adj == null || adj.IsDefault) return BitmapHelper.Clone32bppArgb(source);

            int width = source.Width;
            int height = source.Height;
            var result = BitmapHelper.Clone32bppArgb(source);

            // 1. Color and tone adjustments (Brightness, Contrast, Saturation, Exposure, Opacity)
            bool needsTone = adj.Brightness != 0 || adj.Contrast != 0 || adj.Saturation != 0 || adj.Exposure != 0 || adj.Opacity != 100;
            if (needsTone)
            {
                ApplyToneAdjustments(result, adj);
            }

            // 2. Blur (box blur)
            if (adj.Blur > 0)
            {
                ApplyFastBlur(result, Math.Min(20, adj.Blur));
            }

            // 3. Sharpness (Laplacian convolution)
            if (adj.Sharpness > 0)
            {
                ApplySharpen(result, adj.Sharpness);
            }

            return result;
        }

        private static void ApplyToneAdjustments(Bitmap bitmap, AdjustmentSettings adj)
        {
            float brightFactor = 1.0f + (adj.Brightness / 100.0f);
            float contrastFactor = (100.0f + adj.Contrast) / 100.0f;
            contrastFactor *= contrastFactor; // quadratic response for smoother feel
            float satFactor = 1.0f + (adj.Saturation / 100.0f);
            float exposureMult = (float)Math.Pow(2.0, adj.Exposure / 50.0);
            float totalBright = brightFactor * exposureMult;
            float opacityFactor = adj.Opacity / 100.0f;

            var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

            try
            {
                unsafe
                {
                    byte* ptr = (byte*)data.Scan0;
                    int stride = data.Stride;

                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        byte* row = ptr + (y * stride);
                        for (int x = 0; x < bitmap.Width; x++)
                        {
                            int px = x * 4;
                            float b = row[px + 0];
                            float g = row[px + 1];
                            float r = row[px + 2];
                            float a = row[px + 3];

                            // Exposure & Brightness
                            r *= totalBright;
                            g *= totalBright;
                            b *= totalBright;

                            // Contrast: (pixel/255 - 0.5) * contrast + 0.5
                            if (adj.Contrast != 0)
                            {
                                r = (((r / 255.0f - 0.5f) * contrastFactor) + 0.5f) * 255.0f;
                                g = (((g / 255.0f - 0.5f) * contrastFactor) + 0.5f) * 255.0f;
                                b = (((b / 255.0f - 0.5f) * contrastFactor) + 0.5f) * 255.0f;
                            }

                            // Saturation: lerp between luminance and color
                            if (adj.Saturation != 0)
                            {
                                float lum = 0.299f * r + 0.587f * g + 0.114f * b;
                                r = lum + (r - lum) * satFactor;
                                g = lum + (g - lum) * satFactor;
                                b = lum + (b - lum) * satFactor;
                            }

                            // Opacity
                            a *= opacityFactor;

                            row[px + 0] = (byte)Math.Min(255, Math.Max(0, (int)Math.Round(b)));
                            row[px + 1] = (byte)Math.Min(255, Math.Max(0, (int)Math.Round(g)));
                            row[px + 2] = (byte)Math.Min(255, Math.Max(0, (int)Math.Round(r)));
                            row[px + 3] = (byte)Math.Min(255, Math.Max(0, (int)Math.Round(a)));
                        }
                    }
                }
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }

        private static void ApplySharpen(Bitmap bitmap, int amount)
        {
            if (amount <= 0) return;
            float strength = (amount / 100.0f) * 1.5f;

            int width = bitmap.Width;
            int height = bitmap.Height;
            var copy = BitmapHelper.Clone32bppArgb(bitmap);

            var rect = new Rectangle(0, 0, width, height);
            var srcData = copy.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var dstData = bitmap.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);

            try
            {
                unsafe
                {
                    byte* src = (byte*)srcData.Scan0;
                    byte* dst = (byte*)dstData.Scan0;
                    int stride = srcData.Stride;

                    float centerWeight = 1.0f + 4.0f * strength;
                    float edgeWeight = -strength;

                    for (int y = 1; y < height - 1; y++)
                    {
                        byte* srcRow = src + (y * stride);
                        byte* dstRow = dst + (y * stride);
                        byte* upRow = src + ((y - 1) * stride);
                        byte* downRow = src + ((y + 1) * stride);

                        for (int x = 1; x < width - 1; x++)
                        {
                            int px = x * 4;
                            for (int c = 0; c < 3; c++) // B, G, R
                            {
                                float center = srcRow[px + c];
                                float up = upRow[px + c];
                                float down = downRow[px + c];
                                float left = srcRow[px - 4 + c];
                                float right = srcRow[px + 4 + c];

                                float val = center * centerWeight + (up + down + left + right) * edgeWeight;
                                dstRow[px + c] = (byte)Math.Min(255, Math.Max(0, (int)Math.Round(val)));
                            }
                            dstRow[px + 3] = srcRow[px + 3]; // preserve Alpha
                        }
                    }
                }
            }
            finally
            {
                copy.UnlockBits(srcData);
                bitmap.UnlockBits(dstData);
                copy.Dispose();
            }
        }

        private static void ApplyFastBlur(Bitmap bitmap, int radius)
        {
            if (radius <= 0) return;

            int width = bitmap.Width;
            int height = bitmap.Height;
            var rect = new Rectangle(0, 0, width, height);
            var data = bitmap.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

            try
            {
                unsafe
                {
                    byte* ptr = (byte*)data.Scan0;
                    int stride = data.Stride;

                    // Horizontal blur pass
                    byte[] temp = new byte[width * height * 4];
                    for (int y = 0; y < height; y++)
                    {
                        byte* row = ptr + (y * stride);
                        int rowOffset = y * width * 4;

                        for (int x = 0; x < width; x++)
                        {
                            int sumB = 0, sumG = 0, sumR = 0, sumA = 0;
                            int count = 0;

                            for (int dx = -radius; dx <= radius; dx++)
                            {
                                int nx = x + dx;
                                if (nx >= 0 && nx < width)
                                {
                                    int npx = nx * 4;
                                    sumB += row[npx + 0];
                                    sumG += row[npx + 1];
                                    sumR += row[npx + 2];
                                    sumA += row[npx + 3];
                                    count++;
                                }
                            }

                            int dstPx = rowOffset + (x * 4);
                            temp[dstPx + 0] = (byte)(sumB / count);
                            temp[dstPx + 1] = (byte)(sumG / count);
                            temp[dstPx + 2] = (byte)(sumR / count);
                            temp[dstPx + 3] = (byte)(sumA / count);
                        }
                    }

                    // Vertical blur pass
                    for (int x = 0; x < width; x++)
                    {
                        for (int y = 0; y < height; y++)
                        {
                            int sumB = 0, sumG = 0, sumR = 0, sumA = 0;
                            int count = 0;

                            for (int dy = -radius; dy <= radius; dy++)
                            {
                                int ny = y + dy;
                                if (ny >= 0 && ny < height)
                                {
                                    int nIdx = (ny * width + x) * 4;
                                    sumB += temp[nIdx + 0];
                                    sumG += temp[nIdx + 1];
                                    sumR += temp[nIdx + 2];
                                    sumA += temp[nIdx + 3];
                                    count++;
                                }
                            }

                            byte* targetPixel = ptr + (y * stride) + (x * 4);
                            targetPixel[0] = (byte)(sumB / count);
                            targetPixel[1] = (byte)(sumG / count);
                            targetPixel[2] = (byte)(sumR / count);
                            targetPixel[3] = (byte)(sumA / count);
                        }
                    }
                }
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }
    }
}

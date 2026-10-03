using System;
using System.Drawing;
using System.Drawing.Imaging;

namespace OneAssist.ImageEditor.Core
{
    /// <summary>
    /// Represents a 1-byte per pixel alpha/selection mask (0 = masked/background, 255 = foreground).
    /// </summary>
    public class ImageMask
    {
        public int Width { get; }
        public int Height { get; }
        public byte[] Data { get; }

        public ImageMask(int width, int height, byte initialValue = 255)
        {
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));

            Width = width;
            Height = height;
            Data = new byte[width * height];
            if (initialValue != 0)
            {
                for (int i = 0; i < Data.Length; i++) Data[i] = initialValue;
            }
        }

        public ImageMask(int width, int height, byte[] existingData)
        {
            Width = width;
            Height = height;
            Data = (byte[])existingData.Clone();
        }

        public byte GetPixel(int x, int y)
        {
            if (x < 0 || x >= Width || y < 0 || y >= Height) return 0;
            return Data[y * Width + x];
        }

        public void SetPixel(int x, int y, byte value)
        {
            if (x >= 0 && x < Width && y >= 0 && y < Height)
            {
                Data[y * Width + x] = value;
            }
        }

        public void Clear(byte value)
        {
            for (int i = 0; i < Data.Length; i++)
            {
                Data[i] = value;
            }
        }

        public void Invert()
        {
            for (int i = 0; i < Data.Length; i++)
            {
                Data[i] = (byte)(255 - Data[i]);
            }
        }

        public ImageMask Clone()
        {
            return new ImageMask(Width, Height, Data);
        }

        /// <summary>
        /// Paints a circle onto the mask with hardness and opacity.
        /// </summary>
        public void PaintCircle(int cx, int cy, int radius, byte targetValue, float hardness = 1.0f, float opacity = 1.0f)
        {
            if (radius <= 0) return;

            int minX = Math.Max(0, cx - radius);
            int maxX = Math.Min(Width - 1, cx + radius);
            int minY = Math.Max(0, cy - radius);
            int maxY = Math.Min(Height - 1, cy + radius);

            float rSq = radius * radius;
            float innerRadius = radius * Math.Max(0f, Math.Min(1f, hardness));
            float innerSq = innerRadius * innerRadius;
            float falloffRange = Math.Max(1f, radius - innerRadius);

            for (int y = minY; y <= maxY; y++)
            {
                int dy = y - cy;
                int dySq = dy * dy;
                int rowOffset = y * Width;

                for (int x = minX; x <= maxX; x++)
                {
                    int dx = x - cx;
                    int distSq = dx * dx + dySq;

                    if (distSq <= rSq)
                    {
                        float alphaFactor = 1.0f;
                        if (distSq > innerSq)
                        {
                            float dist = (float)Math.Sqrt(distSq);
                            alphaFactor = 1.0f - ((dist - innerRadius) / falloffRange);
                            alphaFactor = Math.Max(0f, Math.Min(1f, alphaFactor));
                        }

                        float effectiveAlpha = alphaFactor * opacity;
                        int idx = rowOffset + x;
                        byte current = Data[idx];
                        byte blended = (byte)Math.Round(current + (targetValue - current) * effectiveAlpha);
                        Data[idx] = blended;
                    }
                }
            }
        }

        /// <summary>
        /// Paints a continuous line between two points with circle brush.
        /// </summary>
        public void PaintLine(int x0, int y0, int x1, int y1, int radius, byte targetValue, float hardness = 1.0f, float opacity = 1.0f)
        {
            int dx = Math.Abs(x1 - x0);
            int dy = Math.Abs(y1 - y0);
            int distance = (int)Math.Ceiling(Math.Sqrt(dx * dx + dy * dy));

            // Step size at most 1/4 brush radius or 1px
            int stepSize = Math.Max(1, radius / 4);
            int steps = Math.Max(1, distance / stepSize);

            for (int i = 0; i <= steps; i++)
            {
                float t = (float)i / steps;
                int x = (int)Math.Round(x0 + (x1 - x0) * t);
                int y = (int)Math.Round(y0 + (y1 - y0) * t);
                PaintCircle(x, y, radius, targetValue, hardness, opacity);
            }
        }

        /// <summary>
        /// Extracts alpha channel from 32bpp ARGB Bitmap into this mask.
        /// </summary>
        public static ImageMask FromBitmapAlpha(Bitmap bitmap)
        {
            var mask = new ImageMask(bitmap.Width, bitmap.Height);
            var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

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
                            // In 32bpp ARGB, byte 3 is Alpha (B=0, G=1, R=2, A=3)
                            mask.Data[y * bitmap.Width + x] = row[x * 4 + 3];
                        }
                    }
                }
            }
            finally
            {
                bitmap.UnlockBits(data);
            }

            return mask;
        }

        /// <summary>
        /// Applies this alpha mask to a 32bpp ARGB bitmap.
        /// </summary>
        public void ApplyToBitmap(Bitmap bitmap)
        {
            if (bitmap.Width != Width || bitmap.Height != Height) return;

            var rect = new Rectangle(0, 0, Width, Height);
            var data = bitmap.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

            try
            {
                unsafe
                {
                    byte* ptr = (byte*)data.Scan0;
                    int stride = data.Stride;

                    for (int y = 0; y < Height; y++)
                    {
                        byte* row = ptr + (y * stride);
                        int maskOffset = y * Width;

                        for (int x = 0; x < Width; x++)
                        {
                            row[x * 4 + 3] = Data[maskOffset + x];
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

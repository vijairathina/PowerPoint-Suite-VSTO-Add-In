using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

namespace OneAssist.ImageEditor.Utils
{
    /// <summary>
    /// High-performance bitmap utilities for GDI+ and WPF interoperability.
    /// </summary>
    public static class BitmapHelper
    {
        /// <summary>
        /// Creates a deep copy of a Bitmap guaranteed to be 32bpp ARGB.
        /// </summary>
        public static Bitmap Clone32bppArgb(Bitmap source)
        {
            if (source == null) return null;
            var clone = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(clone))
            {
                g.CompositingMode = CompositingMode.SourceCopy;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(source, 0, 0, source.Width, source.Height);
            }
            return clone;
        }

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        /// <summary>
        /// Converts GDI+ Bitmap to WPF BitmapSource for UI display at maximum speed without memory leaks.
        /// </summary>
        public static BitmapSource ToBitmapSource(Bitmap bitmap)
        {
            if (bitmap == null) return null;

            IntPtr hBitmap = IntPtr.Zero;
            try
            {
                hBitmap = bitmap.GetHbitmap();
                var bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

                bitmapSource.Freeze(); // Freezing allows WPF cross-thread access and hardware acceleration
                return bitmapSource;
            }
            finally
            {
                if (hBitmap != IntPtr.Zero)
                {
                    DeleteObject(hBitmap);
                }
            }
        }

        /// <summary>
        /// Converts WPF BitmapSource back to GDI+ Bitmap.
        /// </summary>
        public static Bitmap ToGdiBitmap(BitmapSource source)
        {
            if (source == null) return null;

            using (var outStream = new MemoryStream())
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(source));
                encoder.Save(outStream);
                outStream.Position = 0;
                return new Bitmap(outStream);
            }
        }

        /// <summary>
        /// Loads a Bitmap from file safely into memory, closing the file handle immediately.
        /// </summary>
        public static Bitmap LoadBitmapFromDiskSafely(string filePath)
        {
            if (!File.Exists(filePath)) return null;

            byte[] bytes = File.ReadAllBytes(filePath);
            using (var ms = new MemoryStream(bytes))
            {
                using (var temp = Image.FromStream(ms))
                {
                    return Clone32bppArgb(new Bitmap(temp));
                }
            }
        }

        /// <summary>
        /// Saves a bitmap losslessly to a PNG byte array.
        /// </summary>
        public static byte[] ToPngBytes(Bitmap bitmap)
        {
            if (bitmap == null) return null;
            using (var ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Png);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Rotates/flips a bitmap and returns a new transformed 32bpp ARGB bitmap.
        /// </summary>
        public static Bitmap TransformBitmap(Bitmap source, RotateFlipType rotateFlipType)
        {
            if (source == null) return null;
            var clone = Clone32bppArgb(source);
            clone.RotateFlip(rotateFlipType);
            return clone;
        }

        /// <summary>
        /// Crops a bitmap safely within boundaries.
        /// </summary>
        public static Bitmap CropBitmap(Bitmap source, Rectangle cropRect)
        {
            if (source == null) return null;

            // Constrain rect within source image boundaries
            int x = Math.Max(0, Math.Min(cropRect.X, source.Width - 1));
            int y = Math.Max(0, Math.Min(cropRect.Y, source.Height - 1));
            int w = Math.Max(1, Math.Min(cropRect.Width, source.Width - x));
            int h = Math.Max(1, Math.Min(cropRect.Height, source.Height - y));

            var target = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(target))
            {
                g.CompositingMode = CompositingMode.SourceCopy;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(source, new Rectangle(0, 0, w, h), new Rectangle(x, y, w, h), GraphicsUnit.Pixel);
            }
            return target;
        }
    }
}

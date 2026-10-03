using System;
using System.Drawing;
using System.Drawing.Imaging;
using OneAssist.ImageEditor.Utils;

namespace OneAssist.ImageEditor.Core
{
    /// <summary>
    /// Coordinates the multi-layer image document: Original, Working, AlphaMask, RemovalMask, Background, and Adjustments.
    /// </summary>
    public class ImageDocument : IDisposable
    {
        public Bitmap OriginalImage { get; private set; }
        public Bitmap WorkingImage { get; private set; }
        public ImageMask AlphaMask { get; private set; }
        public ImageMask RemovalMask { get; private set; }
        public AdjustmentSettings Adjustments { get; } = new AdjustmentSettings();
        public BackgroundSettings Background { get; } = new BackgroundSettings();
        public ShapeMetadata Metadata { get; set; } = new ShapeMetadata();
        public ImageHistory History { get; } = new ImageHistory();

        public bool HasImage => WorkingImage != null;
        public int Width => WorkingImage?.Width ?? 0;
        public int Height => WorkingImage?.Height ?? 0;

        public event EventHandler DocumentChanged;

        public void LoadNewImage(Bitmap source, ShapeMetadata metadata)
        {
            DisposeBitmaps();
            History.Clear();
            Adjustments.Reset();
            Background.Type = BackgroundType.Transparent;

            OriginalImage = BitmapHelper.Clone32bppArgb(source);
            WorkingImage = BitmapHelper.Clone32bppArgb(source);
            AlphaMask = ImageMask.FromBitmapAlpha(WorkingImage);
            RemovalMask = new ImageMask(WorkingImage.Width, WorkingImage.Height, 0);
            Metadata = metadata ?? new ShapeMetadata();

            NotifyChanged();
        }

        public void ResetToOriginal()
        {
            if (OriginalImage == null) return;

            History.Push("Reset", WorkingImage, AlphaMask);

            WorkingImage?.Dispose();
            WorkingImage = BitmapHelper.Clone32bppArgb(OriginalImage);
            AlphaMask = ImageMask.FromBitmapAlpha(WorkingImage);
            RemovalMask = new ImageMask(WorkingImage.Width, WorkingImage.Height, 0);
            Adjustments.Reset();
            Background.Type = BackgroundType.Transparent;

            NotifyChanged();
        }

        public void UpdateWorkingImage(Bitmap newWorkingImage, string historyDescription = null)
        {
            if (newWorkingImage == null) return;

            if (!string.IsNullOrEmpty(historyDescription))
            {
                History.Push(historyDescription, WorkingImage, AlphaMask);
            }

            WorkingImage?.Dispose();
            WorkingImage = BitmapHelper.Clone32bppArgb(newWorkingImage);

            if (AlphaMask == null || AlphaMask.Width != WorkingImage.Width || AlphaMask.Height != WorkingImage.Height)
            {
                AlphaMask = ImageMask.FromBitmapAlpha(WorkingImage);
            }

            if (RemovalMask == null || RemovalMask.Width != WorkingImage.Width || RemovalMask.Height != WorkingImage.Height)
            {
                RemovalMask = new ImageMask(WorkingImage.Width, WorkingImage.Height, 0);
            }

            NotifyChanged();
        }

        public void UpdateMask(ImageMask newMask, string historyDescription = null)
        {
            if (newMask == null) return;

            if (!string.IsNullOrEmpty(historyDescription))
            {
                History.Push(historyDescription, WorkingImage, AlphaMask);
            }

            AlphaMask = newMask.Clone();
            NotifyChanged();
        }

        public void ClearRemovalMask()
        {
            RemovalMask?.Clear(0);
            NotifyChanged();
        }

        public void RestoreLineFromOriginal(int x0, int y0, int x1, int y1, int radius, float hardness, float opacity)
        {
            if (OriginalImage == null || WorkingImage == null || AlphaMask == null) return;

            int dx = Math.Abs(x1 - x0);
            int dy = Math.Abs(y1 - y0);
            int distance = (int)Math.Ceiling(Math.Sqrt(dx * dx + dy * dy));
            int stepSize = Math.Max(1, radius / 4);
            int steps = Math.Max(1, distance / stepSize);

            for (int i = 0; i <= steps; i++)
            {
                float t = (float)i / steps;
                int x = (int)Math.Round(x0 + (x1 - x0) * t);
                int y = (int)Math.Round(y0 + (y1 - y0) * t);
                RestoreCircleFromOriginal(x, y, radius, hardness, opacity);
            }
        }

        public void RestoreCircleFromOriginal(int cx, int cy, int radius, float hardness, float opacity)
        {
            if (OriginalImage == null || WorkingImage == null || AlphaMask == null) return;
            if (OriginalImage.Width != WorkingImage.Width || OriginalImage.Height != WorkingImage.Height) return;

            int w = WorkingImage.Width;
            int h = WorkingImage.Height;

            // 1. Restore AlphaMask
            AlphaMask.PaintCircle(cx, cy, radius, 255, hardness, opacity);

            // 2. Restore RGB pixels from pristine OriginalImage
            int minX = Math.Max(0, cx - radius);
            int maxX = Math.Min(w - 1, cx + radius);
            int minY = Math.Max(0, cy - radius);
            int maxY = Math.Min(h - 1, cy + radius);

            if (minX > maxX || minY > maxY) return;

            var rect = new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
            var origData = OriginalImage.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var workData = WorkingImage.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

            try
            {
                unsafe
                {
                    byte* origPtr = (byte*)origData.Scan0;
                    byte* workPtr = (byte*)workData.Scan0;
                    int origStride = origData.Stride;
                    int workStride = workData.Stride;

                    float rSq = radius * radius;
                    float innerRadius = radius * Math.Max(0f, Math.Min(1f, hardness));
                    float innerSq = innerRadius * innerRadius;
                    float falloff = Math.Max(1f, radius - innerRadius);

                    for (int y = minY; y <= maxY; y++)
                    {
                        int localY = y - minY;
                        byte* origRow = origPtr + (localY * origStride);
                        byte* workRow = workPtr + (localY * workStride);
                        int dy = y - cy;
                        int dySq = dy * dy;

                        for (int x = minX; x <= maxX; x++)
                        {
                            int localX = x - minX;
                            int dx = x - cx;
                            int distSq = dx * dx + dySq;

                            if (distSq <= rSq)
                            {
                                float factor = 1.0f;
                                if (distSq > innerSq)
                                {
                                    float d = (float)Math.Sqrt(distSq);
                                    factor = 1.0f - ((d - innerRadius) / falloff);
                                    factor = Math.Max(0f, Math.Min(1f, factor));
                                }
                                float effAlpha = factor * opacity;
                                int px = localX * 4;

                                for (int c = 0; c < 4; c++)
                                {
                                    byte origVal = origRow[px + c];
                                    byte workVal = workRow[px + c];
                                    workRow[px + c] = (byte)Math.Round(workVal + (origVal - workVal) * effAlpha);
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                OriginalImage.UnlockBits(origData);
                WorkingImage.UnlockBits(workData);
            }
        }

        public void Undo()
        {
            if (!History.CanUndo) return;

            var restoredStep = History.PopUndo(WorkingImage, AlphaMask, "Undo");
            if (restoredStep != null)
            {
                WorkingImage?.Dispose();
                WorkingImage = BitmapHelper.Clone32bppArgb(restoredStep.WorkingImage);
                AlphaMask = restoredStep.Mask?.Clone();
                restoredStep.Dispose();
                NotifyChanged();
            }
        }

        public void Redo()
        {
            if (!History.CanRedo) return;

            var restoredStep = History.PopRedo(WorkingImage, AlphaMask, "Redo");
            if (restoredStep != null)
            {
                WorkingImage?.Dispose();
                WorkingImage = BitmapHelper.Clone32bppArgb(restoredStep.WorkingImage);
                AlphaMask = restoredStep.Mask?.Clone();
                restoredStep.Dispose();
                NotifyChanged();
            }
        }

        public void NotifyChanged()
        {
            DocumentChanged?.Invoke(this, EventArgs.Empty);
        }

        private void DisposeBitmaps()
        {
            OriginalImage?.Dispose();
            OriginalImage = null;
            WorkingImage?.Dispose();
            WorkingImage = null;
            AlphaMask = null;
            RemovalMask = null;
        }

        public void Dispose()
        {
            DisposeBitmaps();
            Background.BackgroundImage?.Dispose();
            Background.BackgroundImage = null;
            History.Dispose();
        }
    }
}

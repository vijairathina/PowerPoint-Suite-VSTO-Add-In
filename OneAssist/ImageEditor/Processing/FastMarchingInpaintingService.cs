using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;
using System.Threading.Tasks;
using OneAssist.ImageEditor.Core;
using OneAssist.ImageEditor.Utils;

namespace OneAssist.ImageEditor.Processing
{
    /// <summary>
    /// Object removal inpainting using fast marching inward color diffusion.
    /// </summary>
    public class FastMarchingInpaintingService : IObjectRemovalService
    {
        public Task<Bitmap> RemoveObjectAsync(Bitmap image, Bitmap mask)
        {
            var imageMask = mask != null ? ImageMask.FromBitmapAlpha(mask) : null;
            return RemoveObjectAsync(image, imageMask, new InpaintOptions());
        }

        public async Task<Bitmap> RemoveObjectAsync(
            Bitmap image,
            ImageMask mask,
            InpaintOptions options = null,
            IProgress<int> progress = null,
            CancellationToken cancellationToken = default)
        {
            if (image == null) return null;
            if (mask == null) return BitmapHelper.Clone32bppArgb(image);

            options = options ?? new InpaintOptions();

            return await Task.Run(() =>
            {
                int width = image.Width;
                int height = image.Height;
                var result = BitmapHelper.Clone32bppArgb(image);

                progress?.Report(10);
                cancellationToken.ThrowIfCancellationRequested();

                var rect = new Rectangle(0, 0, width, height);
                var data = result.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

                try
                {
                    unsafe
                    {
                        byte* ptr = (byte*)data.Scan0;
                        int stride = data.Stride;

                        // 0 = known, 1 = band (boundary), 2 = inside mask (to fill)
                        byte[] status = new byte[width * height];
                        int totalToFill = 0;

                        for (int i = 0; i < status.Length; i++)
                        {
                            // In mask, marked area has value > 20
                            if (mask.Data[i] > 20)
                            {
                                status[i] = 2; // to fill
                                totalToFill++;
                            }
                            else
                            {
                                status[i] = 0; // known background
                            }
                        }

                        if (totalToFill == 0)
                        {
                            return result;
                        }

                        progress?.Report(20);
                        cancellationToken.ThrowIfCancellationRequested();

                        // Find boundary band pixels (to-fill pixels adjacent to known pixels)
                        var queue = new Queue<int>(totalToFill / 2 + 100);
                        for (int y = 0; y < height; y++)
                        {
                            for (int x = 0; x < width; x++)
                            {
                                int idx = y * width + x;
                                if (status[idx] == 2)
                                {
                                    bool hasKnown = false;
                                    if (x > 0 && status[idx - 1] == 0) hasKnown = true;
                                    else if (x < width - 1 && status[idx + 1] == 0) hasKnown = true;
                                    else if (y > 0 && status[idx - width] == 0) hasKnown = true;
                                    else if (y < height - 1 && status[idx + width] == 0) hasKnown = true;

                                    if (hasKnown)
                                    {
                                        status[idx] = 1; // band
                                        queue.Enqueue(idx);
                                    }
                                }
                            }
                        }

                        int radius = Math.Max(3, options.Radius);
                        int filledCount = 0;
                        int lastProgress = 20;

                        // Inward fast marching propagation
                        while (queue.Count > 0)
                        {
                            int currentIdx = queue.Dequeue();
                            int cx = currentIdx % width;
                            int cy = currentIdx / width;

                            float totalWeight = 0f;
                            float sumR = 0f;
                            float sumG = 0f;
                            float sumB = 0f;

                            int rMin = Math.Max(0, cy - radius);
                            int rMax = Math.Min(height - 1, cy + radius);
                            int cMin = Math.Max(0, cx - radius);
                            int cMax = Math.Min(width - 1, cx + radius);

                            for (int ny = rMin; ny <= rMax; ny++)
                            {
                                for (int nx = cMin; nx <= cMax; nx++)
                                {
                                    int nIdx = ny * width + nx;
                                    if (status[nIdx] == 0) // known pixel
                                    {
                                        int dx = nx - cx;
                                        int dy = ny - cy;
                                        float distSq = dx * dx + dy * dy;
                                        if (distSq == 0) continue;

                                        float weight = 1.0f / (float)Math.Sqrt(distSq);
                                        byte* nPixel = ptr + (ny * stride) + (nx * 4);

                                        sumB += nPixel[0] * weight;
                                        sumG += nPixel[1] * weight;
                                        sumR += nPixel[2] * weight;
                                        totalWeight += weight;
                                    }
                                }
                            }

                            byte* cPixel = ptr + (cy * stride) + (cx * 4);
                            if (totalWeight > 0f)
                            {
                                cPixel[0] = (byte)Math.Min(255, Math.Max(0, sumB / totalWeight));
                                cPixel[1] = (byte)Math.Min(255, Math.Max(0, sumG / totalWeight));
                                cPixel[2] = (byte)Math.Min(255, Math.Max(0, sumR / totalWeight));
                            }

                            // Mark as known
                            status[currentIdx] = 0;
                            filledCount++;

                            // Enqueue new neighbor pixels inside mask
                            if (cx > 0 && status[currentIdx - 1] == 2)
                            {
                                status[currentIdx - 1] = 1;
                                queue.Enqueue(currentIdx - 1);
                            }
                            if (cx < width - 1 && status[currentIdx + 1] == 2)
                            {
                                status[currentIdx + 1] = 1;
                                queue.Enqueue(currentIdx + 1);
                            }
                            if (cy > 0 && status[currentIdx - width] == 2)
                            {
                                status[currentIdx - width] = 1;
                                queue.Enqueue(currentIdx - width);
                            }
                            if (cy < height - 1 && status[currentIdx + width] == 2)
                            {
                                status[currentIdx + width] = 1;
                                queue.Enqueue(currentIdx + width);
                            }

                            if (filledCount % 2000 == 0)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                int currentPercent = 20 + (int)((float)filledCount / totalToFill * 75);
                                if (currentPercent > lastProgress)
                                {
                                    lastProgress = currentPercent;
                                    progress?.Report(currentPercent);
                                }
                            }
                        }
                    }
                }
                finally
                {
                    result.UnlockBits(data);
                }

                progress?.Report(100);
                return result;
            }, cancellationToken);
        }
    }
}

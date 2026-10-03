using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading;
using System.Threading.Tasks;
using OneAssist.ImageEditor.Utils;

namespace OneAssist.ImageEditor.Processing
{
    /// <summary>
    /// Smart background removal using a two-pass GrabCut-inspired approach.
    ///
    /// Key safety mechanisms that prevent removing the foreground subject:
    ///  - Center-of-mass protection: the central 50% of the image is never seeded as background.
    ///  - Minimum foreground requirement: if &gt; 65% of pixels would be removed, aborts and retries
    ///    with half the tolerance; if still too aggressive, returns the image unchanged.
    ///  - Strict perimeter-only seeding: seeds come only from the outermost 2 pixel rows/columns.
    ///  - Adaptive moving-average cluster model for better background color grouping.
    /// </summary>
    public class SmartBackgroundRemovalService : IBackgroundRemovalService
    {
        public Task<Bitmap> RemoveBackgroundAsync(Bitmap image)
        {
            return RemoveBackgroundAsync(image, new BackgroundRemovalOptions());
        }

        public async Task<Bitmap> RemoveBackgroundAsync(
            Bitmap image,
            BackgroundRemovalOptions options,
            IProgress<int> progress = null,
            CancellationToken cancellationToken = default)
        {
            if (image == null) return null;
            options = options ?? new BackgroundRemovalOptions();

            return await Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();

                int width = image.Width;
                int height = image.Height;
                var result = BitmapHelper.Clone32bppArgb(image);

                progress?.Report(5);

                var rect = new Rectangle(0, 0, width, height);
                var data = result.LockBits(rect, ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);

                try
                {
                    unsafe
                    {
                        byte* ptr = (byte*)data.Scan0;
                        int stride = data.Stride;

                        // Map slider 10..90 → perceptual distance 25..120
                        float baseTolerance = 25f + (options.Tolerance - 10f) * (95f / 80f);
                        int feather = Math.Max(0, Math.Min(20, options.Feather));

                        // Step 1: Build background color model from image perimeter
                        progress?.Report(10);
                        var bgClusters = BuildBackgroundColorModel(ptr, stride, width, height);

                        progress?.Report(20);
                        cancellationToken.ThrowIfCancellationRequested();

                        // Step 2: Flood fill from border with safety checks
                        byte[] bgMask;
                        bool safeToApply = TryFloodFillBackground(
                            ptr, stride, width, height,
                            bgClusters, baseTolerance,
                            progress, cancellationToken,
                            out bgMask);

                        progress?.Report(70);
                        cancellationToken.ThrowIfCancellationRequested();

                        // If too aggressive, retry with reduced tolerance
                        if (!safeToApply)
                        {
                            float reducedTolerance = baseTolerance * 0.45f;
                            safeToApply = TryFloodFillBackground(
                                ptr, stride, width, height,
                                bgClusters, reducedTolerance,
                                null, cancellationToken,
                                out bgMask);
                        }

                        // Still removing too much? Return image unchanged.
                        if (!safeToApply)
                        {
                            return result;
                        }

                        progress?.Report(75);
                        cancellationToken.ThrowIfCancellationRequested();

                        // Step 3: Build feathered alpha array
                        byte[] finalAlpha = BuildFeatheredAlpha(bgMask, width, height, feather);

                        progress?.Report(90);

                        // Step 4: Write alpha to result bitmap
                        for (int y = 0; y < height; y++)
                        {
                            byte* row = ptr + (y * stride);
                            int maskOffset = y * width;
                            for (int x = 0; x < width; x++)
                            {
                                row[x * 4 + 3] = finalAlpha[maskOffset + x];
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

        // Build a moving-average cluster list from the image perimeter (outermost 2 pixel border)
        private static unsafe List<Color3> BuildBackgroundColorModel(byte* ptr, int stride, int width, int height)
        {
            var samples = new List<Color3>(200);
            int stepX = Math.Max(1, width / 60);
            int stepY = Math.Max(1, height / 60);

            for (int x = 0; x < width; x += stepX)
            {
                samples.Add(GetPixelColor(ptr, stride, x, 0));
                samples.Add(GetPixelColor(ptr, stride, x, height - 1));
                if (height > 4)
                {
                    samples.Add(GetPixelColor(ptr, stride, x, 1));
                    samples.Add(GetPixelColor(ptr, stride, x, height - 2));
                }
            }
            for (int y = 1; y < height - 1; y += stepY)
            {
                samples.Add(GetPixelColor(ptr, stride, 0, y));
                samples.Add(GetPixelColor(ptr, stride, width - 1, y));
                if (width > 4)
                {
                    samples.Add(GetPixelColor(ptr, stride, 1, y));
                    samples.Add(GetPixelColor(ptr, stride, width - 2, y));
                }
            }

            // Group into clusters using moving-average
            const float clusterRadius = 28f;
            const int maxClusters = 30;
            var clusters = new List<Color3>(maxClusters);

            foreach (var s in samples)
            {
                bool matched = false;
                for (int i = 0; i < clusters.Count; i++)
                {
                    if (ColorDistance(s, clusters[i]) < clusterRadius)
                    {
                        clusters[i] = new Color3(
                            (byte)((clusters[i].R + s.R) / 2),
                            (byte)((clusters[i].G + s.G) / 2),
                            (byte)((clusters[i].B + s.B) / 2));
                        matched = true;
                        break;
                    }
                }
                if (!matched && clusters.Count < maxClusters)
                    clusters.Add(s);
            }

            return clusters;
        }

        // BFS flood fill from the image border.
        // Returns false when more than 65% of pixels would be removed (unsafe – would eat the foreground).
        private static unsafe bool TryFloodFillBackground(
            byte* ptr, int stride, int width, int height,
            List<Color3> bgClusters, float tolerance,
            IProgress<int> progress,
            CancellationToken cancellationToken,
            out byte[] bgMask)
        {
            int total = width * height;
            bgMask = new byte[total]; // 0 = foreground (keep), 1 = background (remove)

            byte[] visited = new byte[total];
            var queue = new Queue<int>(Math.Min(total, width * 8));

            // Safe interior zone – central 50% of image is NEVER seeded as background
            int safeX1 = width / 4;
            int safeX2 = width - width / 4;
            int safeY1 = height / 4;
            int safeY2 = height - height / 4;

            // Seed from all 4 image borders
            for (int x = 0; x < width; x++)
            {
                TryEnqueueSeed(ptr, stride, width, x, 0, bgClusters, tolerance, visited, queue);
                TryEnqueueSeed(ptr, stride, width, x, height - 1, bgClusters, tolerance, visited, queue);
            }
            for (int y = 1; y < height - 1; y++)
            {
                TryEnqueueSeed(ptr, stride, width, 0, y, bgClusters, tolerance, visited, queue);
                TryEnqueueSeed(ptr, stride, width, width - 1, y, bgClusters, tolerance, visited, queue);
            }

            int bgCount = 0;
            int processed = 0;

            while (queue.Count > 0)
            {
                int idx = queue.Dequeue();
                int px = idx % width;
                int py = idx / width;

                // Never classify safe interior zone as background
                if (px >= safeX1 && px < safeX2 && py >= safeY1 && py < safeY2)
                    continue;

                bgMask[idx] = 1;
                bgCount++;
                processed++;

                if (processed % 40000 == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    progress?.Report(20 + Math.Min(45, (int)((float)processed / total * 45)));
                }

                if (px > 0) CheckNeighbor(ptr, stride, width, idx - 1, px - 1, py, bgClusters, tolerance, visited, queue);
                if (px < width - 1) CheckNeighbor(ptr, stride, width, idx + 1, px + 1, py, bgClusters, tolerance, visited, queue);
                if (py > 0) CheckNeighbor(ptr, stride, width, idx - width, px, py - 1, bgClusters, tolerance, visited, queue);
                if (py < height - 1) CheckNeighbor(ptr, stride, width, idx + width, px, py + 1, bgClusters, tolerance, visited, queue);
            }

            // If > 65% would be removed, the algorithm is being too aggressive
            float bgFraction = (float)bgCount / total;
            return bgFraction <= 0.65f;
        }

        // Build alpha with soft feathering along the foreground/background boundary
        private static byte[] BuildFeatheredAlpha(byte[] bgMask, int width, int height, int feather)
        {
            byte[] alpha = new byte[width * height];
            for (int i = 0; i < alpha.Length; i++)
                alpha[i] = (byte)(bgMask[i] == 1 ? 0 : 255);

            if (feather <= 0) return alpha;

            byte[] result = new byte[alpha.Length];

            for (int y = 0; y < height; y++)
            {
                int rowOff = y * width;
                for (int x = 0; x < width; x++)
                {
                    int cur = alpha[rowOff + x];

                    bool isEdge = false;
                    if (x > 0 && alpha[rowOff + x - 1] != cur) isEdge = true;
                    else if (x < width - 1 && alpha[rowOff + x + 1] != cur) isEdge = true;
                    else if (y > 0 && alpha[rowOff - width + x] != cur) isEdge = true;
                    else if (y < height - 1 && alpha[rowOff + width + x] != cur) isEdge = true;

                    if (!isEdge)
                    {
                        result[rowOff + x] = (byte)cur;
                        continue;
                    }

                    long sum = 0;
                    int count = 0;
                    for (int dy = -feather; dy <= feather; dy++)
                    {
                        int ny = y + dy;
                        if (ny < 0 || ny >= height) continue;
                        int nRowOff = ny * width;
                        for (int dx = -feather; dx <= feather; dx++)
                        {
                            int nx = x + dx;
                            if (nx < 0 || nx >= width) continue;
                            sum += alpha[nRowOff + nx];
                            count++;
                        }
                    }
                    result[rowOff + x] = (byte)(sum / count);
                }
            }
            return result;
        }

        private static unsafe Color3 GetPixelColor(byte* ptr, int stride, int x, int y)
        {
            byte* p = ptr + (y * stride) + (x * 4);
            return new Color3(p[2], p[1], p[0]); // R, G, B
        }

        private static unsafe void TryEnqueueSeed(
            byte* ptr, int stride, int width, int x, int y,
            List<Color3> clusters, float tolerance, byte[] visited, Queue<int> queue)
        {
            int idx = y * width + x;
            if (visited[idx] == 1) return;

            Color3 color = GetPixelColor(ptr, stride, x, y);
            float minDist = float.MaxValue;
            foreach (var cluster in clusters)
            {
                float d = ColorDistance(color, cluster);
                if (d < minDist) minDist = d;
            }

            if (minDist <= tolerance)
            {
                visited[idx] = 1;
                queue.Enqueue(idx);
            }
        }

        private static unsafe void CheckNeighbor(
            byte* ptr, int stride, int width, int idx, int x, int y,
            List<Color3> clusters, float tolerance, byte[] visited, Queue<int> queue)
        {
            if (visited[idx] == 1) return;
            visited[idx] = 1;

            Color3 color = GetPixelColor(ptr, stride, x, y);
            float minDist = float.MaxValue;
            foreach (var cluster in clusters)
            {
                float d = ColorDistance(color, cluster);
                if (d < minDist) minDist = d;
            }

            if (minDist <= tolerance)
            {
                queue.Enqueue(idx);
            }
        }

        private static float ColorDistance(Color3 c1, Color3 c2)
        {
            int rmean = (c1.R + c2.R) / 2;
            int dr = c1.R - c2.R;
            int dg = c1.G - c2.G;
            int db = c1.B - c2.B;

            return (float)Math.Sqrt(
                (((512 + rmean) * dr * dr) >> 8) +
                4 * dg * dg +
                (((767 - rmean) * db * db) >> 8)
            );
        }

        private struct Color3
        {
            public byte R;
            public byte G;
            public byte B;

            public Color3(byte r, byte g, byte b)
            {
                R = r;
                G = g;
                B = b;
            }
        }
    }
}

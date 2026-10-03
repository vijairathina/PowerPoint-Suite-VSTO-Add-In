using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using OneAssist.ImageEditor.Core;

namespace OneAssist.ImageEditor.Processing
{
    public class InpaintOptions
    {
        public int Radius { get; set; } = 8;
        public int Iterations { get; set; } = 1;
    }

    /// <summary>
    /// Service abstraction for object removal/inpainting.
    /// Can be implemented by local fast marching, PatchMatch, or AI models.
    /// </summary>
    public interface IObjectRemovalService
    {
        Task<Bitmap> RemoveObjectAsync(
            Bitmap image,
            Bitmap mask);

        Task<Bitmap> RemoveObjectAsync(
            Bitmap image,
            ImageMask mask,
            InpaintOptions options = null,
            IProgress<int> progress = null,
            CancellationToken cancellationToken = default);
    }
}

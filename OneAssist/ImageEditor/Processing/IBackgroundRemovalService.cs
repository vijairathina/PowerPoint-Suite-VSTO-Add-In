using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace OneAssist.ImageEditor.Processing
{
    public class BackgroundRemovalOptions
    {
        public int Tolerance { get; set; } = 38; // 10 to 100
        public int Feather { get; set; } = 3;    // 0 to 20
        public int Smoothness { get; set; } = 2; // 1 to 5
    }

    /// <summary>
    /// Service abstraction for AI/algorithmic background removal.
    /// Can be implemented by local algorithms, ONNX Runtime, or cloud APIs.
    /// </summary>
    public interface IBackgroundRemovalService
    {
        Task<Bitmap> RemoveBackgroundAsync(Bitmap image);

        Task<Bitmap> RemoveBackgroundAsync(
            Bitmap image, 
            BackgroundRemovalOptions options, 
            IProgress<int> progress = null, 
            CancellationToken cancellationToken = default);
    }
}

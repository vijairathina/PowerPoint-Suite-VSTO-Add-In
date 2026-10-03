using System;
using System.Drawing;
using OneAssist.ImageEditor.Utils;

namespace OneAssist.ImageEditor.Core
{
    public class HistoryStep : IDisposable
    {
        public string Description { get; }
        public DateTime Timestamp { get; }
        public Bitmap WorkingImage { get; }
        public ImageMask Mask { get; }

        public HistoryStep(string description, Bitmap workingImage, ImageMask mask)
        {
            Description = description;
            Timestamp = DateTime.UtcNow;
            WorkingImage = workingImage != null ? BitmapHelper.Clone32bppArgb(workingImage) : null;
            Mask = mask?.Clone();
        }

        public void Dispose()
        {
            WorkingImage?.Dispose();
        }
    }
}

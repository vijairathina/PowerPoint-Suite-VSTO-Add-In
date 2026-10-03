using System;
using System.IO;
using System.Threading;
using OneAssist.SlideCompare.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.SlideCompare.Extraction
{
    /// <summary>
    /// Extracts a complete PresentationSnapshot from a PowerPoint Presentation COM object.
    /// </summary>
    public static class PresentationExtractor
    {
        public static PresentationSnapshot ExtractPresentation(
            PowerPoint.Presentation presentation,
            IProgress<Tuple<int, string>> progress = null,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (presentation == null)
                throw new ArgumentNullException(nameof(presentation));

            var snapshot = new PresentationSnapshot();

            try { snapshot.PresentationName = presentation.Name ?? "Presentation"; } catch { snapshot.PresentationName = "Presentation"; }
            try { snapshot.FilePath = presentation.FullName ?? string.Empty; } catch { }
            try { snapshot.SlideWidth = presentation.PageSetup.SlideWidth; } catch { snapshot.SlideWidth = 960f; }
            try { snapshot.SlideHeight = presentation.PageSetup.SlideHeight; } catch { snapshot.SlideHeight = 540f; }

            int slideCount = 0;
            try { slideCount = presentation.Slides.Count; } catch { }

            for (int i = 1; i <= slideCount; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                int percent = (int)((i / (double)Math.Max(1, slideCount)) * 100);
                progress?.Report(Tuple.Create(percent, $"Extracting slide {i} of {slideCount}..."));

                try
                {
                    var slide = presentation.Slides[i];
                    var slideSnapshot = SlideExtractor.ExtractSlide(slide, i);
                    if (slideSnapshot != null)
                    {
                        snapshot.Slides.Add(slideSnapshot);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Warning: Failed to extract slide {i}: {ex.Message}");
                }
            }

            return snapshot;
        }
    }
}

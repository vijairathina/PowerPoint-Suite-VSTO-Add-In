using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using OneAssist.ImageEditor.Core;
using OneAssist.ImageEditor.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.ImageEditor.PowerPointServices
{
    public class ExtractedImageResult
    {
        public Bitmap Image { get; set; }
        public ShapeMetadata Metadata { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// Handles detecting selected PowerPoint pictures and extracting their images into memory.
    /// Supports standard pictures, linked pictures, placeholders, SVG graphics, and shapes.
    /// </summary>
    public static class PowerPointImageService
    {
        private static PowerPoint.Application ResolveApp(PowerPoint.Application app)
        {
            if (app != null) return app;
            try
            {
                return Globals.ThisAddIn?.Application;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Checks if the current PowerPoint selection contains a shape/picture.
        /// </summary>
        public static bool IsPictureSelected(PowerPoint.Application app, out string reason)
        {
            reason = string.Empty;
            app = ResolveApp(app);

            if (app == null)
            {
                reason = "PowerPoint application is not available.";
                return false;
            }

            try
            {
                PowerPoint.DocumentWindow window = null;
                try { window = app.ActiveWindow; } catch { }
                if (window == null)
                {
                    reason = "No active presentation window found.";
                    return false;
                }

                PowerPoint.Selection selection = null;
                try { selection = window.Selection; } catch { }
                if (selection == null)
                {
                    reason = "Please select an image on your slide.";
                    return false;
                }

                // Check selection type
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    reason = "Please select an image in PowerPoint.";
                    return false;
                }

                PowerPoint.ShapeRange shapeRange = null;
                try { shapeRange = selection.ShapeRange; } catch { }
                if (shapeRange == null || shapeRange.Count < 1)
                {
                    reason = "Please select an image in PowerPoint.";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                reason = $"Selection check: {ex.Message}";
                return false;
            }
        }

        /// <summary>
        /// Extracts the selected picture from PowerPoint into a 32bpp ARGB Bitmap without leaving permanent temp files.
        /// Uses high-resolution PNG export with automatic Clipboard fallback.
        /// </summary>
        public static ExtractedImageResult ExtractSelectedImage(PowerPoint.Application app)
        {
            var result = new ExtractedImageResult();
            app = ResolveApp(app);

            if (app == null)
            {
                result.Success = false;
                result.ErrorMessage = "PowerPoint application is not available.";
                return result;
            }

            PowerPoint.DocumentWindow window = null;
            PowerPoint.Selection selection = null;
            PowerPoint.ShapeRange shapeRange = null;
            PowerPoint.Shape shape = null;
            PowerPoint.Slide slide = null;
            string tempFile = null;

            try
            {
                try { window = app.ActiveWindow; } catch { }
                if (window == null)
                {
                    result.Success = false;
                    result.ErrorMessage = "No active presentation window found.";
                    return result;
                }

                try { selection = window.Selection; } catch { }
                if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    result.Success = false;
                    result.ErrorMessage = "Please select an image in PowerPoint.";
                    return result;
                }

                try { shapeRange = selection.ShapeRange; } catch { }
                if (shapeRange == null || shapeRange.Count < 1)
                {
                    result.Success = false;
                    result.ErrorMessage = "Please select an image in PowerPoint.";
                    return result;
                }

                shape = shapeRange[1];

                // Safely determine slide index
                int slideIndex = 1;
                try
                {
                    if (window.View != null && window.View.Slide != null)
                    {
                        slide = (PowerPoint.Slide)window.View.Slide;
                        slideIndex = slide.SlideIndex;
                    }
                }
                catch { }

                if (slide == null)
                {
                    try
                    {
                        slide = shape.Parent as PowerPoint.Slide;
                        if (slide != null) slideIndex = slide.SlideIndex;
                    }
                    catch { }
                }

                if (slide == null)
                {
                    try
                    {
                        var pres = app.ActivePresentation;
                        if (pres != null && pres.Slides.Count > 0)
                        {
                            slide = pres.Slides[1];
                            slideIndex = 1;
                        }
                    }
                    catch { }
                }

                var meta = new ShapeMetadata
                {
                    Left = shape.Left,
                    Top = shape.Top,
                    Width = shape.Width,
                    Height = shape.Height,
                    Rotation = shape.Rotation,
                    ZOrderPosition = shape.ZOrderPosition,
                    Name = shape.Name,
                    AlternativeText = string.Empty,
                    Title = string.Empty,
                    ShapeId = shape.Id,
                    SlideIndex = slideIndex,
                    IsValid = true
                };

                try { meta.AlternativeText = shape.AlternativeText; } catch { }
                try { meta.Title = shape.Title; } catch { }

                // 1. Try Exporting to high-resolution PNG
                Bitmap loadedBitmap = null;
                using (var tempMgr = new TempFileManager())
                {
                    tempFile = tempMgr.GetTempPngPath();

                    int exportW = Math.Max(400, (int)Math.Round(shape.Width * 2.5f));
                    int exportH = Math.Max(400, (int)Math.Round(shape.Height * 2.5f));

                    bool exportOk = false;
                    try
                    {
                        shape.Export(tempFile, PowerPoint.PpShapeFormat.ppShapeFormatPNG, exportW, exportH, PowerPoint.PpExportMode.ppScaleToFit);
                        exportOk = File.Exists(tempFile) && new FileInfo(tempFile).Length > 0;
                    }
                    catch { }

                    if (!exportOk)
                    {
                        try
                        {
                            shape.Export(tempFile, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 0, 0, PowerPoint.PpExportMode.ppScaleToFit);
                            exportOk = File.Exists(tempFile) && new FileInfo(tempFile).Length > 0;
                        }
                        catch { }
                    }

                    if (exportOk)
                    {
                        loadedBitmap = BitmapHelper.LoadBitmapFromDiskSafely(tempFile);
                    }
                }

                // 2. Fallback to Clipboard Copy if Export did not produce a bitmap
                if (loadedBitmap == null)
                {
                    try
                    {
                        shape.Copy();
                        Thread.Sleep(60);

                        for (int retry = 0; retry < 3; retry++)
                        {
                            if (System.Windows.Forms.Clipboard.ContainsImage())
                            {
                                using (var clipImg = System.Windows.Forms.Clipboard.GetImage())
                                {
                                    if (clipImg != null)
                                    {
                                        loadedBitmap = BitmapHelper.Clone32bppArgb(new Bitmap(clipImg));
                                        break;
                                    }
                                }
                            }
                            Thread.Sleep(50);
                        }
                    }
                    catch { }
                }

                if (loadedBitmap != null)
                {
                    result.Image = loadedBitmap;
                    result.Metadata = meta;
                    result.Success = true;
                    return result;
                }

                result.Success = false;
                result.ErrorMessage = "Unable to extract image data from the selected shape.";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error extracting image: {ex.Message}";
                return result;
            }
            finally
            {
                TempFileManager.SafeDelete(tempFile);

                if (shape != null) Marshal.ReleaseComObject(shape);
                if (shapeRange != null) Marshal.ReleaseComObject(shapeRange);
                if (selection != null) Marshal.ReleaseComObject(selection);
                if (slide != null) Marshal.ReleaseComObject(slide);
                if (window != null) Marshal.ReleaseComObject(window);
            }
        }
    }
}

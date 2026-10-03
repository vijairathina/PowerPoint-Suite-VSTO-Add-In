using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using OneAssist.ImageEditor.Core;
using OneAssist.ImageEditor.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.ImageEditor.PowerPointServices
{
    /// <summary>
    /// Applies the edited image back into PowerPoint, preserving shape dimensions, rotation, position, and metadata.
    /// </summary>
    public static class PowerPointApplyService
    {
        public static bool ApplyImageToPowerPoint(
            PowerPoint.Application app,
            Bitmap editedImage,
            ShapeMetadata metadata,
            out string errorMessage)
        {
            errorMessage = string.Empty;
            if (app == null)
            {
                errorMessage = "PowerPoint application is unavailable.";
                return false;
            }

            if (editedImage == null)
            {
                errorMessage = "No edited image available to apply.";
                return false;
            }

            if (metadata == null || !metadata.IsValid)
            {
                errorMessage = "Original shape metadata is invalid.";
                return false;
            }

            PowerPoint.Presentation presentation = null;
            PowerPoint.Slide slide = null;
            PowerPoint.Shape oldShape = null;
            PowerPoint.Shape newShape = null;
            string tempFile = null;

            try
            {
                presentation = app.ActivePresentation;
                if (presentation == null)
                {
                    errorMessage = "No active presentation found.";
                    return false;
                }

                // Locate target slide
                try
                {
                    slide = presentation.Slides[metadata.SlideIndex];
                }
                catch
                {
                    // Fallback to active slide
                    var window = app.ActiveWindow;
                    if (window?.View?.Slide != null)
                    {
                        slide = (PowerPoint.Slide)window.View.Slide;
                    }
                }

                if (slide == null)
                {
                    errorMessage = "Could not locate the target slide.";
                    return false;
                }

                // Locate original shape
                try
                {
                    foreach (PowerPoint.Shape s in slide.Shapes)
                    {
                        if (s.Id == metadata.ShapeId || s.Name == metadata.Name)
                        {
                            oldShape = s;
                            break;
                        }
                    }
                }
                catch
                {
                    // Continue even if old shape cannot be found by Id
                }

                // Save edited image to scoped temporary PNG
                using (var tempMgr = new TempFileManager())
                {
                    tempFile = tempMgr.GetTempPngPath();
                    editedImage.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);

                    // Add new picture with exact position and dimensions
                    newShape = slide.Shapes.AddPicture(
                        tempFile,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        metadata.Left,
                        metadata.Top,
                        metadata.Width,
                        metadata.Height);

                    if (newShape != null)
                    {
                        // Restore attributes
                        newShape.Rotation = metadata.Rotation;
                        newShape.Name = metadata.Name;
                        if (!string.IsNullOrEmpty(metadata.AlternativeText))
                        {
                            newShape.AlternativeText = metadata.AlternativeText;
                        }
                        if (!string.IsNullOrEmpty(metadata.Title))
                        {
                            newShape.Title = metadata.Title;
                        }

                        // Restore Z-Order if old shape exists
                        if (oldShape != null)
                        {
                            int targetZ = metadata.ZOrderPosition;
                            // Match ZOrder
                            while (newShape.ZOrderPosition > targetZ && targetZ > 1)
                            {
                                int currentZ = newShape.ZOrderPosition;
                                newShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                                if (newShape.ZOrderPosition == currentZ) break; // Cannot move further
                            }

                            // Delete the old shape now that new one is in place
                            try
                            {
                                oldShape.Delete();
                            }
                            catch
                            {
                                // Ignore error if old shape was already removed
                            }
                        }

                        // Update metadata shape ID to new shape
                        metadata.ShapeId = newShape.Id;

                        // Select the newly added picture
                        try
                        {
                            newShape.Select();
                        }
                        catch { }

                        return true;
                    }
                }

                errorMessage = "Failed to create picture on the slide.";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = $"Error applying image to PowerPoint: {ex.Message}";
                return false;
            }
            finally
            {
                TempFileManager.SafeDelete(tempFile);

                if (newShape != null) Marshal.ReleaseComObject(newShape);
                if (oldShape != null) Marshal.ReleaseComObject(oldShape);
                if (slide != null) Marshal.ReleaseComObject(slide);
                if (presentation != null) Marshal.ReleaseComObject(presentation);
            }
        }
    }
}

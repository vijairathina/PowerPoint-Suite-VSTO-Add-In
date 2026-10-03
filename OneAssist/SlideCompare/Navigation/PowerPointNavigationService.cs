using System;
using System.Linq;
using OneAssist.SlideCompare.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneAssist.SlideCompare.Navigation
{
    /// <summary>
    /// Handles navigating PowerPoint to the destination slide and selecting the target shape.
    /// </summary>
    public static class PowerPointNavigationService
    {
        public static bool NavigateToShape(
            PowerPoint.Presentation destinationPresentation,
            int slideNumber,
            string shapeName,
            int shapeId = 0,
            int shapeIndex = 0)
        {
            if (destinationPresentation == null) return false;

            try
            {
                // 1. Ensure destination presentation is active
                try
                {
                    destinationPresentation.Windows[1].Activate();
                }
                catch { }

                // 2. Locate slide
                if (slideNumber <= 0 || slideNumber > destinationPresentation.Slides.Count)
                    return false;

                PowerPoint.Slide targetSlide = destinationPresentation.Slides[slideNumber];
                targetSlide.Select();

                try
                {
                    var window = destinationPresentation.Windows[1];
                    window.View.GotoSlide(slideNumber);
                }
                catch { }

                // 3. Locate shape on the slide
                PowerPoint.Shape targetShape = null;

                // Try finding by ID
                if (shapeId > 0)
                {
                    foreach (PowerPoint.Shape shp in targetSlide.Shapes)
                    {
                        try
                        {
                            if (shp.Id == shapeId)
                            {
                                targetShape = shp;
                                break;
                            }
                        }
                        catch { }
                    }
                }

                // Try finding by Name
                if (targetShape == null && !string.IsNullOrEmpty(shapeName))
                {
                    foreach (PowerPoint.Shape shp in targetSlide.Shapes)
                    {
                        try
                        {
                            if (string.Equals(shp.Name, shapeName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetShape = shp;
                                break;
                            }
                        }
                        catch { }
                    }
                }

                // Try finding by Index
                if (targetShape == null && shapeIndex > 0 && shapeIndex <= targetSlide.Shapes.Count)
                {
                    try
                    {
                        targetShape = targetSlide.Shapes[shapeIndex];
                    }
                    catch { }
                }

                // 4. Select shape
                if (targetShape != null)
                {
                    try
                    {
                        targetShape.Select(Office.MsoTriState.msoTrue);
                    }
                    catch { }
                    return true;
                }

                return true; // Navigated to slide even if specific shape wasn't selectable
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error navigating to shape: {ex.Message}");
                return false;
            }
        }

        public static bool NavigateToSlide(PowerPoint.Presentation destinationPresentation, int slideNumber)
        {
            if (destinationPresentation == null || slideNumber <= 0) return false;

            try
            {
                try { destinationPresentation.Windows[1].Activate(); } catch { }

                if (slideNumber <= destinationPresentation.Slides.Count)
                {
                    PowerPoint.Slide targetSlide = destinationPresentation.Slides[slideNumber];
                    targetSlide.Select();
                    try
                    {
                        destinationPresentation.Windows[1].View.GotoSlide(slideNumber);
                    }
                    catch { }
                    return true;
                }
            }
            catch { }

            return false;
        }
    }
}

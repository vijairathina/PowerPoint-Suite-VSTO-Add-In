using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public class ImageQualityIssue
    {
        public int ShapeId { get; set; }
        public string ShapeName { get; set; }
        public string Problem { get; set; }
        public string Recommendation { get; set; }
        public bool IsDistorted { get; set; }
    }

    public static class ImageQualityEngine
    {
        public static List<ImageQualityIssue> InspectShapeImage(PowerPoint.Shape shape)
        {
            var issues = new List<ImageQualityIssue>();
            if (shape == null) return issues;

            try
            {
                if (shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture)
                {
                    // Check aspect ratio distortion
                    // If shape has scale factor difference between width and height
                    try
                    {
                        // In PowerPoint, if LockAspectRatio is msoFalse and Width/Height are heavily skewed
                        if (shape.LockAspectRatio == Office.MsoTriState.msoFalse)
                        {
                            // If shape is stretched unnaturally (> 40% aspect distortion)
                            float ratio = shape.Width / Math.Max(1.0f, shape.Height);
                            if (ratio > 10.0f || ratio < 0.1f)
                            {
                                issues.Add(new ImageQualityIssue
                                {
                                    ShapeId = shape.Id,
                                    ShapeName = shape.Name,
                                    Problem = "Image aspect ratio appears heavily distorted or stretched.",
                                    Recommendation = "Reset picture to original aspect ratio to preserve visual fidelity.",
                                    IsDistorted = true
                                });
                            }
                        }
                    }
                    catch { }

                    // Check for oversized dimensions
                    if (shape.Width > 2000 || shape.Height > 2000)
                    {
                        issues.Add(new ImageQualityIssue
                        {
                            ShapeId = shape.Id,
                            ShapeName = shape.Name,
                            Problem = $"Image dimension ({shape.Width:F0}x{shape.Height:F0}pt) is exceptionally large and exceeds standard slide canvas.",
                            Recommendation = "Compress or scale down picture to improve presentation performance.",
                            IsDistorted = false
                        });
                    }
                }
            }
            catch { }

            return issues;
        }

        public static void ResetAspectRatio(PowerPoint.Shape shape)
        {
            if (shape == null) return;
            try
            {
                shape.ScaleWidth(1.0f, Office.MsoTriState.msoTrue);
                shape.ScaleHeight(1.0f, Office.MsoTriState.msoTrue);
                shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            }
            catch { }
        }
    }
}

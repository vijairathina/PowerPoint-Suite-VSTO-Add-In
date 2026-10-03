using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public class CleanupCandidate
    {
        public int ShapeId { get; set; }
        public string ShapeName { get; set; }
        public string Reason { get; set; }
        public int SlideIndex { get; set; }
    }

    public static class SlideCleanupService
    {
        public static List<CleanupCandidate> FindCleanupCandidates(PowerPoint.Slide slide, int slideIndex)
        {
            var candidates = new List<CleanupCandidate>();
            if (slide == null) return candidates;

            try
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        // 1. Zero size shapes
                        if (shape.Width <= 1.0f && shape.Height <= 1.0f)
                        {
                            candidates.Add(new CleanupCandidate
                            {
                                ShapeId = shape.Id,
                                ShapeName = shape.Name,
                                Reason = "Zero-size invisible artifact (Width/Height <= 1pt).",
                                SlideIndex = slideIndex
                            });
                            continue;
                        }

                        // 2. Empty text boxes
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            var tf = shape.TextFrame;
                            if (tf.HasText == Office.MsoTriState.msoTrue)
                            {
                                string txt = tf.TextRange.Text;
                                if (string.IsNullOrWhiteSpace(txt))
                                {
                                    // Shape has text frame with only whitespace and no visible background or border
                                    bool hasFill = shape.Fill != null && shape.Fill.Visible == Office.MsoTriState.msoTrue;
                                    bool hasLine = shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue;
                                    if (!hasFill && !hasLine)
                                    {
                                        candidates.Add(new CleanupCandidate
                                        {
                                            ShapeId = shape.Id,
                                            ShapeName = shape.Name,
                                            Reason = "Empty text box containing only whitespace.",
                                            SlideIndex = slideIndex
                                        });
                                        continue;
                                    }
                                }
                            }
                        }

                        // 3. Invisible shapes
                        if (shape.Visible == Office.MsoTriState.msoFalse)
                        {
                            candidates.Add(new CleanupCandidate
                            {
                                ShapeId = shape.Id,
                                ShapeName = shape.Name,
                                Reason = "Hidden/invisible object.",
                                SlideIndex = slideIndex
                            });
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return candidates;
        }

        public static int DeleteCandidate(PowerPoint.Slide slide, int shapeId)
        {
            if (slide == null) return 0;
            try
            {
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (s.Id == shapeId)
                    {
                        s.Delete();
                        return 1;
                    }
                }
            }
            catch { }
            return 0;
        }
    }
}

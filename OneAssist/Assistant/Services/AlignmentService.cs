using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public static class AlignmentService
    {
        public static void AlignShapes(List<PowerPoint.Shape> shapes, AlignmentAction action, float slideWidth = 0, float slideHeight = 0)
        {
            if (shapes == null || shapes.Count == 0) return;

            try
            {
                if (shapes.Count == 1 && (action == AlignmentAction.CenterOnSlideHorizontally || action == AlignmentAction.CenterOnSlideVertically))
                {
                    var s = shapes[0];
                    if (action == AlignmentAction.CenterOnSlideHorizontally && slideWidth > 0)
                    {
                        s.Left = (slideWidth - s.Width) / 2.0f;
                    }
                    else if (action == AlignmentAction.CenterOnSlideVertically && slideHeight > 0)
                    {
                        s.Top = (slideHeight - s.Height) / 2.0f;
                    }
                    return;
                }

                if (shapes.Count < 2) return;

                switch (action)
                {
                    case AlignmentAction.AlignLeft:
                        float minLeft = shapes.Min(s => s.Left);
                        foreach (var s in shapes) s.Left = minLeft;
                        break;

                    case AlignmentAction.AlignRight:
                        float maxRight = shapes.Max(s => s.Left + s.Width);
                        foreach (var s in shapes) s.Left = maxRight - s.Width;
                        break;

                    case AlignmentAction.AlignTop:
                        float minTop = shapes.Min(s => s.Top);
                        foreach (var s in shapes) s.Top = minTop;
                        break;

                    case AlignmentAction.AlignBottom:
                        float maxBottom = shapes.Max(s => s.Top + s.Height);
                        foreach (var s in shapes) s.Top = maxBottom - s.Height;
                        break;

                    case AlignmentAction.AlignHorizontalCenter:
                        float avgHCenter = shapes.Average(s => s.Left + s.Width / 2.0f);
                        foreach (var s in shapes) s.Left = avgHCenter - s.Width / 2.0f;
                        break;

                    case AlignmentAction.AlignVerticalCenter:
                        float avgVCenter = shapes.Average(s => s.Top + s.Height / 2.0f);
                        foreach (var s in shapes) s.Top = avgVCenter - s.Height / 2.0f;
                        break;

                    case AlignmentAction.DistributeHorizontally:
                        DistributeShapesHorizontally(shapes);
                        break;

                    case AlignmentAction.DistributeVertically:
                        DistributeShapesVertically(shapes);
                        break;

                    case AlignmentAction.MatchWidth:
                        float targetWidth = shapes[0].Width;
                        foreach (var s in shapes) s.Width = targetWidth;
                        break;

                    case AlignmentAction.MatchHeight:
                        float targetHeight = shapes[0].Height;
                        foreach (var s in shapes) s.Height = targetHeight;
                        break;

                    case AlignmentAction.MatchBothSize:
                        float tW = shapes[0].Width;
                        float tH = shapes[0].Height;
                        foreach (var s in shapes)
                        {
                            s.Width = tW;
                            s.Height = tH;
                        }
                        break;
                }
            }
            catch { }
        }

        public static void DistributeShapesHorizontally(List<PowerPoint.Shape> shapes)
        {
            if (shapes == null || shapes.Count < 3) return;
            var sorted = shapes.OrderBy(s => s.Left).ToList();

            float totalSpan = (sorted.Last().Left + sorted.Last().Width) - sorted.First().Left;
            float totalWidths = sorted.Sum(s => s.Width);
            float availableGap = totalSpan - totalWidths;
            float gap = availableGap / (sorted.Count - 1);

            float currentLeft = sorted.First().Left + sorted.First().Width + gap;
            for (int i = 1; i < sorted.Count - 1; i++)
            {
                sorted[i].Left = currentLeft;
                currentLeft += sorted[i].Width + gap;
            }
        }

        public static void DistributeShapesVertically(List<PowerPoint.Shape> shapes)
        {
            if (shapes == null || shapes.Count < 3) return;
            var sorted = shapes.OrderBy(s => s.Top).ToList();

            float totalSpan = (sorted.Last().Top + sorted.Last().Height) - sorted.First().Top;
            float totalHeights = sorted.Sum(s => s.Height);
            float availableGap = totalSpan - totalHeights;
            float gap = availableGap / (sorted.Count - 1);

            float currentTop = sorted.First().Top + sorted.First().Height + gap;
            for (int i = 1; i < sorted.Count - 1; i++)
            {
                sorted[i].Top = currentTop;
                currentTop += sorted[i].Height + gap;
            }
        }

        public static void SnapShapeToGrid(PowerPoint.Shape shape, float gridSize = 10.0f)
        {
            if (shape == null || gridSize <= 0) return;
            try
            {
                shape.Left = (float)Math.Round(shape.Left / gridSize) * gridSize;
                shape.Top = (float)Math.Round(shape.Top / gridSize) * gridSize;
                shape.Width = (float)Math.Round(shape.Width / gridSize) * gridSize;
                shape.Height = (float)Math.Round(shape.Height / gridSize) * gridSize;
            }
            catch { }
        }
    }
}

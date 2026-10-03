using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Grid
{
    public class GridColumn
    {
        public int Index { get; set; }
        public float Left { get; set; }
        public float Width { get; set; }
        public float Right => Left + Width;
    }

    public static class SlideGridService
    {
        public static List<GridColumn> Calculate12ColumnGrid(float slideWidth, float margin = 36.0f, float columnGap = 10.8f)
        {
            var cols = new List<GridColumn>();
            if (slideWidth <= 0) return cols;

            float availableWidth = slideWidth - (margin * 2.0f);
            float totalGaps = columnGap * 11.0f;
            float colWidth = (availableWidth - totalGaps) / 12.0f;

            float curLeft = margin;
            for (int i = 1; i <= 12; i++)
            {
                cols.Add(new GridColumn
                {
                    Index = i,
                    Left = curLeft,
                    Width = colWidth
                });
                curLeft += colWidth + columnGap;
            }

            return cols;
        }

        public static void SnapShapeToGridColumn(PowerPoint.Shape shape, List<GridColumn> columns)
        {
            if (shape == null || columns == null || columns.Count == 0) return;

            try
            {
                // Find nearest column left
                float shapeLeft = shape.Left;
                GridColumn nearest = columns[0];
                float minDiff = Math.Abs(shapeLeft - nearest.Left);

                foreach (var c in columns)
                {
                    float diff = Math.Abs(shapeLeft - c.Left);
                    if (diff < minDiff)
                    {
                        minDiff = diff;
                        nearest = c;
                    }
                }

                shape.Left = nearest.Left;
            }
            catch { }
        }
    }
}

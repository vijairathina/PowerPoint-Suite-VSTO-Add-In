using System;
using System.Drawing;
using OneAssist.Assistant.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.DataVisualization
{
    public static class TableDesignerService
    {
        public static bool ApplyZebraStriping(PowerPoint.Shape tableShape, Color headerColor, Color evenRowColor, Color oddRowColor)
        {
            if (tableShape == null || !tableShape.HasTable.Equals(Office.MsoTriState.msoTrue)) return false;

            try
            {
                var table = tableShape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;

                int headerBgr = SafeInteropHelper.ColorToPowerPointRgb(headerColor);
                int evenBgr = SafeInteropHelper.ColorToPowerPointRgb(evenRowColor);
                int oddBgr = SafeInteropHelper.ColorToPowerPointRgb(oddRowColor);

                for (int r = 1; r <= rows; r++)
                {
                    int targetBgr = (r == 1) ? headerBgr : ((r % 2 == 0) ? evenBgr : oddBgr);

                    for (int c = 1; c <= cols; c++)
                    {
                        var cell = table.Cell(r, c);
                        cell.Shape.Fill.Solid();
                        cell.Shape.Fill.ForeColor.RGB = targetBgr;

                        if (r == 1)
                        {
                            cell.Shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                            cell.Shape.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.White);
                        }
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void SetUniformCellPadding(PowerPoint.Shape tableShape, float padding = 5.0f)
        {
            if (tableShape == null || tableShape.HasTable != Office.MsoTriState.msoTrue) return;

            try
            {
                var table = tableShape.Table;
                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        var cell = table.Cell(r, c);
                        var tf = cell.Shape.TextFrame;
                        tf.MarginLeft = padding;
                        tf.MarginRight = padding;
                        tf.MarginTop = padding;
                        tf.MarginBottom = padding;
                    }
                }
            }
            catch { }
        }
    }
}

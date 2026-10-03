using System;
using System.Collections.Generic;
using OneAssist.SlideCompare.Matching;
using OneAssist.SlideCompare.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneAssist.SlideCompare.Extraction
{
    /// <summary>
    /// Safely extracts table cell contents and geometry from PowerPoint tables.
    /// </summary>
    public static class TableExtractor
    {
        public static bool ExtractTable(PowerPoint.Shape shape, out int rowCount, out int colCount, out List<TableCellSnapshot> cells)
        {
            rowCount = 0;
            colCount = 0;
            cells = new List<TableCellSnapshot>();

            if (shape == null) return false;

            try
            {
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    var table = shape.Table;
                    rowCount = table.Rows.Count;
                    colCount = table.Columns.Count;

                    for (int r = 1; r <= rowCount; r++)
                    {
                        for (int c = 1; c <= colCount; c++)
                        {
                            try
                            {
                                var cell = table.Cell(r, c);
                                string cellText = string.Empty;
                                if (cell.Shape.HasTextFrame == Office.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                {
                                    cellText = cell.Shape.TextFrame.TextRange.Text ?? string.Empty;
                                }

                                cells.Add(new TableCellSnapshot
                                {
                                    RowIndex = r - 1,
                                    ColumnIndex = c - 1,
                                    Text = cellText,
                                    NormalizedText = SimilarityCalculator.NormalizeText(cellText)
                                });
                            }
                            catch { }
                        }
                    }
                    return true;
                }
            }
            catch { }

            return false;
        }
    }
}

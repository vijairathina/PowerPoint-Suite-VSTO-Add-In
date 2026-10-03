using System;
using System.Collections.Generic;
using OneAssist.SlideCompare.Matching;
using OneAssist.SlideCompare.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneAssist.SlideCompare.Extraction
{
    /// <summary>
    /// Recursively extracts shapes, groups, tables, charts, and media into POCO snapshots.
    /// </summary>
    public static class ShapeExtractor
    {
        public static ShapeSnapshot ExtractShape(PowerPoint.Shape shape, int shapeIndex)
        {
            if (shape == null) return null;

            var snapshot = new ShapeSnapshot
            {
                Index = shapeIndex
            };

            try { snapshot.Id = shape.Id; } catch { snapshot.Id = shapeIndex; }
            try { snapshot.Name = shape.Name ?? $"Shape {shapeIndex}"; } catch { snapshot.Name = $"Shape {shapeIndex}"; }
            try { snapshot.ShapeType = shape.Type.ToString(); } catch { }

            // Placeholder check
            try
            {
                if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                {
                    snapshot.IsPlaceholder = true;
                    snapshot.PlaceholderType = shape.PlaceholderFormat.Type.ToString();
                }
            }
            catch { }

            // Geometry
            try { snapshot.Left = shape.Left; } catch { }
            try { snapshot.Top = shape.Top; } catch { }
            try { snapshot.Width = shape.Width; } catch { }
            try { snapshot.Height = shape.Height; } catch { }
            try { snapshot.Rotation = shape.Rotation; } catch { }
            try { snapshot.ZOrderPosition = shape.ZOrderPosition; } catch { }

            // Fill & Line Color
            try
            {
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    int rgb = shape.Fill.ForeColor.RGB;
                    snapshot.FillColorHex = $"#{rgb & 0xFF:X2}{(rgb >> 8) & 0xFF:X2}{(rgb >> 16) & 0xFF:X2}";
                }
            }
            catch { }

            try
            {
                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    int rgb = shape.Line.ForeColor.RGB;
                    snapshot.LineColorHex = $"#{rgb & 0xFF:X2}{(rgb >> 8) & 0xFF:X2}{(rgb >> 16) & 0xFF:X2}";
                }
            }
            catch { }

            // Group check & recursive extraction (Req 26)
            try
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    snapshot.IsGroup = true;
                    int childCount = shape.GroupItems.Count;
                    for (int c = 1; c <= childCount; c++)
                    {
                        try
                        {
                            var child = shape.GroupItems[c];
                            var childSnapshot = ExtractShape(child, c);
                            if (childSnapshot != null)
                            {
                                snapshot.ChildShapes.Add(childSnapshot);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            // Text extraction & Typography
            try
            {
                string text = TextExtractor.ExtractText(shape);
                snapshot.RawText = text;
                snapshot.HasText = !string.IsNullOrWhiteSpace(text);
                snapshot.NormalizedText = SimilarityCalculator.NormalizeText(text);

                if (snapshot.HasText)
                {
                    string fFamily, fColor, align;
                    float fSize;
                    bool isBold, isItalic, isUnderline;

                    TextExtractor.ExtractFontProperties(shape, out fFamily, out fSize, out isBold, out isItalic, out isUnderline, out fColor, out align);
                    snapshot.FontFamily = fFamily;
                    snapshot.FontSize = fSize;
                    snapshot.IsBold = isBold;
                    snapshot.IsItalic = isItalic;
                    snapshot.IsUnderline = isUnderline;
                    snapshot.FontColorHex = fColor;
                    snapshot.TextAlignment = align;
                }
            }
            catch { }

            // Table extraction (Req 27)
            try
            {
                int rCount, cCount;
                List<TableCellSnapshot> cells;
                if (TableExtractor.ExtractTable(shape, out rCount, out cCount, out cells))
                {
                    snapshot.IsTable = true;
                    snapshot.TableRowCount = rCount;
                    snapshot.TableColumnCount = cCount;
                    snapshot.TableCells = cells;

                    // Combine table text for search and similarity
                    var tableWords = new List<string>();
                    foreach (var tc in cells)
                    {
                        if (!string.IsNullOrWhiteSpace(tc.Text))
                            tableWords.Add(tc.Text);
                    }
                    if (string.IsNullOrEmpty(snapshot.RawText) && tableWords.Count > 0)
                    {
                        snapshot.RawText = string.Join(" | ", tableWords);
                        snapshot.NormalizedText = SimilarityCalculator.NormalizeText(snapshot.RawText);
                        snapshot.HasText = true;
                    }
                }
            }
            catch { }

            // Chart extraction (Req 5)
            try
            {
                string chartTitle;
                List<string> series, categories;
                if (ChartExtractor.ExtractChart(shape, out chartTitle, out series, out categories))
                {
                    snapshot.IsChart = true;
                    snapshot.ChartTitle = chartTitle;
                    snapshot.ChartSeriesNames = series;
                    snapshot.ChartCategories = categories;

                    if (string.IsNullOrEmpty(snapshot.RawText) && !string.IsNullOrEmpty(chartTitle))
                    {
                        snapshot.RawText = $"Chart: {chartTitle}";
                        snapshot.NormalizedText = SimilarityCalculator.NormalizeText(snapshot.RawText);
                        snapshot.HasText = true;
                    }
                }
            }
            catch { }

            // Image / Media extraction (Req 5)
            try
            {
                if (shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture)
                {
                    snapshot.IsImage = true;
                    snapshot.ImageAspectWidth = snapshot.Width;
                    snapshot.ImageAspectHeight = snapshot.Height;
                    // Compute visual signature hash from dimensions and type
                    snapshot.ImageHash = $"IMG_{Math.Round(snapshot.Width)}x{Math.Round(snapshot.Height)}";
                }
            }
            catch { }

            // SmartArt extraction (Req 28)
            try
            {
                if (shape.HasSmartArt == Office.MsoTriState.msoTrue)
                {
                    snapshot.HasSmartArt = true;
                    // Extracted via child shapes or text frame
                }
            }
            catch { }

            // Generate stable fingerprint
            snapshot.Fingerprint = ContentMatchingEngine.GenerateShapeFingerprint(snapshot);

            return snapshot;
        }
    }
}

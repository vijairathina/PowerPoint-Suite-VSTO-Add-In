using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneAssist.SlideCompare.Extraction
{
    /// <summary>
    /// Safely extracts text and character formatting from PowerPoint text frames.
    /// </summary>
    public static class TextExtractor
    {
        public static string ExtractText(PowerPoint.Shape shape)
        {
            if (shape == null) return string.Empty;

            // 1. Standard TextFrame
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    string t = shape.TextFrame.TextRange.Text;
                    if (!string.IsNullOrEmpty(t)) return t;
                }
            }
            catch { }

            // 2. TextFrame2 (for shapes/placeholders with rich text in modern PowerPoint)
            try
            {
                if (shape.TextFrame2 != null && shape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                {
                    string t = shape.TextFrame2.TextRange.Text;
                    if (!string.IsNullOrEmpty(t)) return t;
                }
            }
            catch { }

            // 3. SmartArt nodes text
            try
            {
                if (shape.HasSmartArt == Office.MsoTriState.msoTrue && shape.SmartArt != null)
                {
                    var nodeTexts = new System.Collections.Generic.List<string>();
                    int count = shape.SmartArt.Nodes.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            var node = shape.SmartArt.Nodes[i];
                            if (node.TextFrame2 != null && node.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                            {
                                string nt = node.TextFrame2.TextRange.Text;
                                if (!string.IsNullOrWhiteSpace(nt))
                                {
                                    nodeTexts.Add(nt.Trim());
                                }
                            }
                        }
                        catch { }
                    }
                    if (nodeTexts.Count > 0)
                    {
                        return string.Join(" | ", nodeTexts);
                    }
                }
            }
            catch { }

            return string.Empty;
        }

        public static void ExtractFontProperties(
            PowerPoint.Shape shape,
            out string fontFamily,
            out float fontSize,
            out bool isBold,
            out bool isItalic,
            out bool isUnderline,
            out string fontColorHex,
            out string alignment)
        {
            fontFamily = string.Empty;
            fontSize = 0;
            isBold = false;
            isItalic = false;
            isUnderline = false;
            fontColorHex = string.Empty;
            alignment = string.Empty;

            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    var tr = shape.TextFrame.TextRange;
                    var font = tr.Font;

                    fontFamily = font.Name ?? string.Empty;
                    fontSize = font.Size;
                    isBold = font.Bold == Office.MsoTriState.msoTrue;
                    isItalic = font.Italic == Office.MsoTriState.msoTrue;
                    isUnderline = font.Underline == Office.MsoTriState.msoTrue;

                    try
                    {
                        int rgb = font.Color.RGB;
                        fontColorHex = $"#{rgb & 0xFF:X2}{(rgb >> 8) & 0xFF:X2}{(rgb >> 16) & 0xFF:X2}";
                    }
                    catch { }

                    try
                    {
                        alignment = tr.ParagraphFormat.Alignment.ToString();
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}

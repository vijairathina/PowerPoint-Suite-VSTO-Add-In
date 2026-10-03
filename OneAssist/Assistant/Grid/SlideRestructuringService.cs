using System;
using System.Collections.Generic;
using System.Drawing;
using OneAssist.Assistant.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Grid
{
    public static class SlideRestructuringService
    {
        public static List<PowerPoint.Shape> ConvertBulletsToCards(PowerPoint.Slide slide, PowerPoint.Shape bulletShape, int columnsCount = 3)
        {
            var createdCards = new List<PowerPoint.Shape>();
            if (slide == null || bulletShape == null || !SafeInteropHelper.HasText(bulletShape)) return createdCards;

            try
            {
                string rawText = bulletShape.TextFrame.TextRange.Text;
                var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) return createdCards;

                float startX = bulletShape.Left;
                float startY = bulletShape.Top;
                float totalWidth = bulletShape.Width;
                float cardHeight = Math.Max(120.0f, bulletShape.Height);

                int cols = Math.Min(columnsCount, lines.Length);
                float gap = 14.4f; // 0.2 in
                float cardWidth = (totalWidth - (gap * (cols - 1))) / cols;

                for (int i = 0; i < cols; i++)
                {
                    float cardX = startX + i * (cardWidth + gap);
                    var card = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle, cardX, startY, cardWidth, cardHeight);

                    // Style card
                    card.Fill.Solid();
                    card.Fill.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(248, 250, 252));
                    card.Line.Visible = Office.MsoTriState.msoTrue;
                    card.Line.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(203, 213, 225));
                    card.Line.Weight = 1.0f;

                    card.TextFrame.TextRange.Text = lines[i].TrimStart('-', '*', '•', ' ');
                    card.TextFrame.TextRange.Font.Name = "Aptos";
                    card.TextFrame.TextRange.Font.Size = 13.0f;
                    card.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(30, 41, 59));
                    card.TextFrame.MarginLeft = 12.0f;
                    card.TextFrame.MarginTop = 12.0f;
                    card.TextFrame.MarginRight = 12.0f;
                    card.TextFrame.MarginBottom = 12.0f;

                    createdCards.Add(card);
                }

                // Delete original bullets shape after transformation
                bulletShape.Delete();
            }
            catch { }

            return createdCards;
        }

        public static PowerPoint.Shape ConvertToCallout(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            if (slide == null || shape == null || !SafeInteropHelper.HasText(shape)) return null;

            try
            {
                string text = shape.TextFrame.TextRange.Text;
                float left = shape.Left;
                float top = shape.Top;
                float width = shape.Width;
                float height = shape.Height;

                var callout = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                callout.Fill.Solid();
                callout.Fill.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(238, 242, 255)); // Soft indigo
                callout.Line.Visible = Office.MsoTriState.msoFalse;

                // Add thick accent left bar
                var accentBar = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, 6.0f, height);
                accentBar.Fill.Solid();
                accentBar.Fill.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(79, 70, 229)); // Indigo 600
                accentBar.Line.Visible = Office.MsoTriState.msoFalse;

                callout.TextFrame.TextRange.Text = text;
                callout.TextFrame.TextRange.Font.Name = "Aptos";
                callout.TextFrame.TextRange.Font.Size = 14.0f;
                callout.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                callout.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(30, 27, 75));
                callout.TextFrame.MarginLeft = 18.0f;

                shape.Delete();
                return callout;
            }
            catch
            {
                return null;
            }
        }
    }
}

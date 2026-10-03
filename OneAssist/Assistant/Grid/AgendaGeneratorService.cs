using System;
using System.Collections.Generic;
using System.Drawing;
using OneAssist.Assistant.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Grid
{
    public static class AgendaGeneratorService
    {
        public static PowerPoint.Slide GenerateAgendaSlide(PowerPoint.Presentation pres)
        {
            if (pres == null) return null;

            try
            {
                var titles = new List<string>();
                for (int i = 1; i <= pres.Slides.Count; i++)
                {
                    var meta = ConsistencyEngine.DetectSlideTitle(pres.Slides[i], i);
                    if (meta != null && !string.IsNullOrWhiteSpace(meta.ShapeName))
                    {
                        string txt = SafeInteropHelper.GetText(pres.Slides[i].Shapes[meta.ShapeId]);
                        if (!string.IsNullOrWhiteSpace(txt) && !titles.Contains(txt) && titles.Count < 8)
                        {
                            titles.Add(txt.Trim());
                        }
                    }
                }

                // Add new slide right after slide 1
                int targetIndex = Math.Min(2, pres.Slides.Count + 1);
                var newSlide = pres.Slides.Add(targetIndex, PowerPoint.PpSlideLayout.ppLayoutBlank);

                float sw = pres.PageSetup.SlideWidth;
                float sh = pres.PageSetup.SlideHeight;

                // Add Agenda Title
                var titleShape = newSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 54.0f, 40.0f, sw - 108.0f, 50.0f);
                titleShape.TextFrame.TextRange.Text = "Table of Contents";
                titleShape.TextFrame.TextRange.Font.Name = "Aptos Display";
                titleShape.TextFrame.TextRange.Font.Size = 28.0f;
                titleShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                titleShape.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(0, 0x20, 0x60));

                // Add agenda items as numbered rows
                float startY = 120.0f;
                float rowHeight = 44.0f;
                for (int i = 0; i < titles.Count; i++)
                {
                    float y = startY + (i * rowHeight);

                    // Number badge
                    var numShape = newSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 54.0f, y + 4.0f, 28.0f, 28.0f);
                    numShape.Fill.Solid();
                    numShape.Fill.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(0, 0x20, 0x60));
                    numShape.Line.Visible = Office.MsoTriState.msoFalse;
                    numShape.TextFrame.TextRange.Text = (i + 1).ToString("D2");
                    numShape.TextFrame.TextRange.Font.Name = "Aptos";
                    numShape.TextFrame.TextRange.Font.Size = 10.0f;
                    numShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    numShape.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.White);

                    // Text
                    var itemShape = newSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 94.0f, y, sw - 160.0f, 36.0f);
                    itemShape.TextFrame.TextRange.Text = titles[i];
                    itemShape.TextFrame.TextRange.Font.Name = "Aptos";
                    itemShape.TextFrame.TextRange.Font.Size = 15.0f;
                    itemShape.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(30, 41, 59));
                }

                newSlide.Select();
                return newSlide;
            }
            catch
            {
                return null;
            }
        }
    }
}

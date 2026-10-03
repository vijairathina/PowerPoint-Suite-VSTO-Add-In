using System;
using System.Drawing;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Inspector
{
    public class ShapeInspectionDetails
    {
        public int ShapeId { get; set; }
        public string Name { get; set; }
        public string ShapeType { get; set; }

        // Position & Geometry (Points & Inches)
        public float XPoints { get; set; }
        public float YPoints { get; set; }
        public float WidthPoints { get; set; }
        public float HeightPoints { get; set; }
        public float Rotation { get; set; }

        public double XInches => Math.Round(XPoints / 72.0, 2);
        public double YInches => Math.Round(YPoints / 72.0, 2);
        public double WidthInches => Math.Round(WidthPoints / 72.0, 2);
        public double HeightInches => Math.Round(HeightPoints / 72.0, 2);

        // Fill & Border
        public bool HasFill { get; set; }
        public string FillHex { get; set; }
        public float Transparency { get; set; }
        public bool HasBorder { get; set; }
        public string BorderHex { get; set; }
        public float BorderWeight { get; set; }

        // Typography
        public bool HasText { get; set; }
        public string TextPreview { get; set; }
        public string FontName { get; set; }
        public float? FontSize { get; set; }
        public string FontColorHex { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public float MarginLeft { get; set; }
        public float MarginTop { get; set; }

        // Layering
        public int ZOrderPosition { get; set; }
        public bool IsGrouped { get; set; }
    }

    public static class ShapeInspectorService
    {
        public static ShapeInspectionDetails Inspect(PowerPoint.Shape shape)
        {
            if (shape == null) return null;

            var d = new ShapeInspectionDetails
            {
                ShapeId = shape.Id,
                Name = shape.Name,
                ShapeType = shape.Type.ToString(),
                XPoints = shape.Left,
                YPoints = shape.Top,
                WidthPoints = shape.Width,
                HeightPoints = shape.Height,
                Rotation = shape.Rotation,
                ZOrderPosition = shape.ZOrderPosition
            };

            try
            {
                // Fill
                if (shape.Fill != null && shape.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    d.HasFill = true;
                    int bgr = shape.Fill.ForeColor.RGB;
                    d.FillHex = Services.SafeInteropHelper.ColorToHex(Services.SafeInteropHelper.ColorFromPowerPointRgb(bgr));
                    d.Transparency = shape.Fill.Transparency;
                }

                // Border
                if (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    d.HasBorder = true;
                    int bgr = shape.Line.ForeColor.RGB;
                    d.BorderHex = Services.SafeInteropHelper.ColorToHex(Services.SafeInteropHelper.ColorFromPowerPointRgb(bgr));
                    d.BorderWeight = shape.Line.Weight;
                }

                // Typography
                if (Services.SafeInteropHelper.HasText(shape))
                {
                    d.HasText = true;
                    var tr = shape.TextFrame.TextRange;
                    d.TextPreview = tr.Text;
                    if (d.TextPreview != null && d.TextPreview.Length > 60)
                    {
                        d.TextPreview = d.TextPreview.Substring(0, 57) + "...";
                    }

                    var font = tr.Font;
                    d.FontName = font.Name;
                    d.FontSize = font.Size;
                    d.IsBold = font.Bold == Office.MsoTriState.msoTrue;
                    d.IsItalic = font.Italic == Office.MsoTriState.msoTrue;
                    d.FontColorHex = Services.SafeInteropHelper.ColorToHex(Services.SafeInteropHelper.ColorFromPowerPointRgb(font.Color.RGB));

                    var tf = shape.TextFrame;
                    d.MarginLeft = tf.MarginLeft;
                    d.MarginTop = tf.MarginTop;
                }
            }
            catch { }

            return d;
        }

        public static void ApplyNumericBounds(PowerPoint.Shape shape, float? xInches, float? yInches, float? widthInches, float? heightInches, float? rotation)
        {
            if (shape == null) return;
            try
            {
                if (xInches.HasValue) shape.Left = xInches.Value * 72.0f;
                if (yInches.HasValue) shape.Top = yInches.Value * 72.0f;
                if (widthInches.HasValue && widthInches.Value > 0) shape.Width = widthInches.Value * 72.0f;
                if (heightInches.HasValue && heightInches.Value > 0) shape.Height = heightInches.Value * 72.0f;
                if (rotation.HasValue) shape.Rotation = rotation.Value;
            }
            catch { }
        }
    }
}

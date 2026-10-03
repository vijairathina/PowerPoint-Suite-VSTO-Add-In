using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public static class DesignerToolboxService
    {
        public static void ChangeZOrder(PowerPoint.Shape shape, Office.MsoZOrderCmd cmd)
        {
            if (shape == null) return;
            try
            {
                shape.ZOrder(cmd);
            }
            catch { }
        }

        public static PowerPoint.Shape GroupShapes(List<PowerPoint.Shape> shapes, PowerPoint.Slide slide)
        {
            if (shapes == null || shapes.Count < 2 || slide == null) return null;
            try
            {
                var names = new List<string>();
                foreach (var s in shapes) names.Add(s.Name);
                var range = slide.Shapes.Range(names.ToArray());
                return range.Group();
            }
            catch
            {
                return null;
            }
        }

        public static void UngroupShape(PowerPoint.Shape shape)
        {
            if (shape == null || shape.Type != Office.MsoShapeType.msoGroup) return;
            try
            {
                shape.Ungroup();
            }
            catch { }
        }
    }
}

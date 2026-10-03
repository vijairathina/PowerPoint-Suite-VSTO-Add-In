using System;
using System.Collections.Generic;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Layers
{
    public class SlideLayerItem
    {
        public int ShapeId { get; set; }
        public string Name { get; set; }
        public string ShapeType { get; set; }
        public int ZOrderPosition { get; set; }
        public bool IsVisible { get; set; }
        public bool IsMorphNamed => Name != null && Name.StartsWith("!!");
    }

    public static class LayerManagerService
    {
        public static List<SlideLayerItem> GetSlideLayers(PowerPoint.Slide slide)
        {
            var layers = new List<SlideLayerItem>();
            if (slide == null) return layers;

            try
            {
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    layers.Add(new SlideLayerItem
                    {
                        ShapeId = s.Id,
                        Name = s.Name,
                        ShapeType = s.Type.ToString(),
                        ZOrderPosition = s.ZOrderPosition,
                        IsVisible = s.Visible == Office.MsoTriState.msoTrue
                    });
                }
            }
            catch { }

            return layers.OrderByDescending(l => l.ZOrderPosition).ToList();
        }

        public static void SetVisibility(PowerPoint.Slide slide, int shapeId, bool visible)
        {
            if (slide == null) return;
            try
            {
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (s.Id == shapeId)
                    {
                        s.Visible = visible ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                        break;
                    }
                }
            }
            catch { }
        }

        public static void RenameShape(PowerPoint.Slide slide, int shapeId, string newName)
        {
            if (slide == null || string.IsNullOrWhiteSpace(newName)) return;
            try
            {
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (s.Id == shapeId)
                    {
                        s.Name = newName.Trim();
                        break;
                    }
                }
            }
            catch { }
        }

        public static void AutoRenameShapes(PowerPoint.Slide slide)
        {
            if (slide == null) return;
            try
            {
                int titleCount = 0;
                int bodyCount = 0;
                int imageCount = 0;
                int chartCount = 0;
                int shapeCount = 0;

                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (s.HasChart == Office.MsoTriState.msoTrue)
                    {
                        s.Name = $"Chart_{++chartCount:D2}";
                    }
                    else if (s.Type == Office.MsoShapeType.msoPicture || s.Type == Office.MsoShapeType.msoLinkedPicture)
                    {
                        s.Name = $"Image_{++imageCount:D2}";
                    }
                    else if (Services.SafeInteropHelper.HasText(s))
                    {
                        float? sz = Services.SafeInteropHelper.GetFontSize(s);
                        if (sz.HasValue && sz.Value >= 20.0f)
                        {
                            s.Name = $"Title_{++titleCount:D2}";
                        }
                        else
                        {
                            s.Name = $"Body_{++bodyCount:D2}";
                        }
                    }
                    else
                    {
                        s.Name = $"Shape_{++shapeCount:D2}";
                    }
                }
            }
            catch { }
        }
    }
}

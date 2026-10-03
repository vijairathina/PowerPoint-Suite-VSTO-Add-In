using System;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public enum PowerPointSelectionKind
    {
        None,
        Slide,
        SingleText,
        SinglePicture,
        SingleShape,
        MultipleShapes
    }

    public class ContextAwareSelectionService
    {
        public static PowerPointSelectionKind DetectSelectionKind(PowerPoint.Selection sel)
        {
            if (sel == null) return PowerPointSelectionKind.None;

            try
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    return PowerPointSelectionKind.Slide;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var sr = sel.ShapeRange;
                    if (sr == null || sr.Count == 0) return PowerPointSelectionKind.None;
                    if (sr.Count > 1) return PowerPointSelectionKind.MultipleShapes;

                    var single = sr[1];
                    if (single.Type == Office.MsoShapeType.msoPicture || single.Type == Office.MsoShapeType.msoLinkedPicture)
                    {
                        return PowerPointSelectionKind.SinglePicture;
                    }

                    if (SafeInteropHelper.HasText(single))
                    {
                        return PowerPointSelectionKind.SingleText;
                    }

                    return PowerPointSelectionKind.SingleShape;
                }
            }
            catch { }

            return PowerPointSelectionKind.None;
        }
    }
}

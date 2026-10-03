using System;
using System.Collections.Generic;
using OneAssist.Assistant.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Selection
{
    public enum SmartSelectMode
    {
        SameFillColor,
        SameFontName,
        SameFontSize,
        SameShapeType,
        SameDimensions,
        SimilarAll
    }

    public static class SmartSelectService
    {
        public static int SelectMatching(PowerPoint.Slide slide, PowerPoint.Shape referenceShape, SmartSelectMode mode)
        {
            if (slide == null || referenceShape == null) return 0;

            var matchingNames = new List<string> { referenceShape.Name };

            try
            {
                var refFill = SafeInteropHelper.GetFillColor(referenceShape);
                var refFont = SafeInteropHelper.GetFontName(referenceShape);
                var refSize = SafeInteropHelper.GetFontSize(referenceShape);
                var refType = referenceShape.Type;
                float refW = referenceShape.Width;
                float refH = referenceShape.Height;

                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (s.Id == referenceShape.Id) continue;

                    bool match = false;

                    switch (mode)
                    {
                        case SmartSelectMode.SameFillColor:
                            var f = SafeInteropHelper.GetFillColor(s);
                            match = refFill.HasValue && f.HasValue && refFill.Value.ToArgb() == f.Value.ToArgb();
                            break;

                        case SmartSelectMode.SameFontName:
                            var fn = SafeInteropHelper.GetFontName(s);
                            match = !string.IsNullOrEmpty(refFont) && string.Equals(refFont, fn, StringComparison.OrdinalIgnoreCase);
                            break;

                        case SmartSelectMode.SameFontSize:
                            var fs = SafeInteropHelper.GetFontSize(s);
                            match = refSize.HasValue && fs.HasValue && Math.Abs(refSize.Value - fs.Value) < 0.2f;
                            break;

                        case SmartSelectMode.SameShapeType:
                            match = s.Type == refType;
                            break;

                        case SmartSelectMode.SameDimensions:
                            match = Math.Abs(s.Width - refW) < 1.0f && Math.Abs(s.Height - refH) < 1.0f;
                            break;

                        case SmartSelectMode.SimilarAll:
                            bool matchType = s.Type == refType;
                            bool matchFont = string.Equals(refFont, SafeInteropHelper.GetFontName(s), StringComparison.OrdinalIgnoreCase);
                            match = matchType || matchFont;
                            break;
                    }

                    if (match)
                    {
                        matchingNames.Add(s.Name);
                    }
                }

                if (matchingNames.Count > 0)
                {
                    slide.Shapes.Range(matchingNames.ToArray()).Select();
                }
            }
            catch { }

            return matchingNames.Count;
        }
    }
}

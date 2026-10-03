using System;
using System.Collections.Generic;
using OneAssist.Assistant.AILayout.Models;

namespace OneAssist.Assistant.AILayout.Engine
{
    public class ConstraintValidationResult
    {
        public bool IsValid { get; set; } = true;
        public List<string> Violations { get; set; } = new List<string>();
        public List<string> CorrectionsApplied { get; set; } = new List<string>();
    }

    public static class LayoutConstraintEngine
    {
        public const float MinMarginLeft = 36f;   // 0.5 in
        public const float MinMarginTop = 36f;    // 0.5 in
        public const float MinMarginRight = 36f;  // 0.5 in
        public const float MinMarginBottom = 36f; // 0.5 in
        public const float MinSpacing = 12f;      // ~0.17 in minimum gap

        public static ConstraintValidationResult ValidateAndCorrect(LayoutSpecification spec, float slideWidth = 960f, float slideHeight = 540f)
        {
            var result = new ConstraintValidationResult();
            if (spec == null || spec.Regions == null || spec.Regions.Count == 0) return result;

            float maxRight = slideWidth - MinMarginRight;
            float maxBottom = slideHeight - MinMarginBottom;

            // 1. Boundary check and clamp
            foreach (var r in spec.Regions)
            {
                // Left margin
                if (r.X < MinMarginLeft)
                {
                    result.Violations.Add($"Region '{r.Name}' left ({r.X:F1}) violated minimum margin of {MinMarginLeft}pt.");
                    r.X = MinMarginLeft;
                    result.CorrectionsApplied.Add($"Clamped '{r.Name}' X to {MinMarginLeft}pt.");
                    result.IsValid = false;
                }

                // Top margin
                if (r.Y < MinMarginTop)
                {
                    result.Violations.Add($"Region '{r.Name}' top ({r.Y:F1}) violated minimum margin of {MinMarginTop}pt.");
                    r.Y = MinMarginTop;
                    result.CorrectionsApplied.Add($"Clamped '{r.Name}' Y to {MinMarginTop}pt.");
                    result.IsValid = false;
                }

                // Right boundary
                if (r.X + r.Width > maxRight)
                {
                    result.Violations.Add($"Region '{r.Name}' right edge ({r.X + r.Width:F1}) exceeded slide right boundary ({maxRight:F1}pt).");
                    r.Width = Math.Max(50f, maxRight - r.X);
                    result.CorrectionsApplied.Add($"Adjusted '{r.Name}' width to {r.Width:F1}pt to fit boundary.");
                    result.IsValid = false;
                }

                // Bottom boundary
                if (r.Y + r.Height > maxBottom)
                {
                    result.Violations.Add($"Region '{r.Name}' bottom edge ({r.Y + r.Height:F1}) exceeded slide bottom boundary ({maxBottom:F1}pt).");
                    r.Height = Math.Max(30f, maxBottom - r.Y);
                    result.CorrectionsApplied.Add($"Adjusted '{r.Name}' height to {r.Height:F1}pt to fit boundary.");
                    result.IsValid = false;
                }
            }

            // 2. Overlap detection & resolution
            for (int i = 0; i < spec.Regions.Count; i++)
            {
                for (int j = i + 1; j < spec.Regions.Count; j++)
                {
                    var r1 = spec.Regions[i];
                    var r2 = spec.Regions[j];

                    // Check if regions overlap
                    bool xOverlap = r1.X < (r2.X + r2.Width) && (r1.X + r1.Width) > r2.X;
                    bool yOverlap = r1.Y < (r2.Y + r2.Height) && (r1.Y + r1.Height) > r2.Y;

                    if (xOverlap && yOverlap)
                    {
                        result.Violations.Add($"Overlap detected between '{r1.Name}' and '{r2.Name}'.");
                        result.IsValid = false;

                        // Separate vertically or horizontally depending on relative positions
                        if (Math.Abs(r1.Y - r2.Y) > Math.Abs(r1.X - r2.X))
                        {
                            // Vertical separation
                            if (r1.Y < r2.Y)
                            {
                                r2.Y = r1.Y + r1.Height + MinSpacing;
                            }
                            else
                            {
                                r1.Y = r2.Y + r2.Height + MinSpacing;
                            }
                            result.CorrectionsApplied.Add($"Separated '{r1.Name}' and '{r2.Name}' vertically by {MinSpacing}pt.");
                        }
                        else
                        {
                            // Horizontal separation
                            if (r1.X < r2.X)
                            {
                                r2.X = r1.X + r1.Width + MinSpacing;
                            }
                            else
                            {
                                r1.X = r2.X + r2.Width + MinSpacing;
                            }
                            result.CorrectionsApplied.Add($"Separated '{r1.Name}' and '{r2.Name}' horizontally by {MinSpacing}pt.");
                        }
                    }
                }
            }

            return result;
        }
    }
}

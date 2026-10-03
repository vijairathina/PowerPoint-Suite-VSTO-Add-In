using System;
using System.Collections.Generic;
using OneAssist.SlideCompare.Models;
using OneAssist.SlideCompare.Services;

namespace OneAssist.SlideCompare.Comparison
{
    /// <summary>
    /// Compares visual and typography formatting between two shapes.
    /// </summary>
    public static class FormattingComparisonEngine
    {
        public static List<FormattingDifferenceItem> CompareFormatting(ShapeSnapshot source, ShapeSnapshot dest, CompareSettings settings)
        {
            var diffs = new List<FormattingDifferenceItem>();
            if (source == null || dest == null || settings == null || !settings.CompareFormatting)
                return diffs;

            // 1. Font Family
            if (!string.IsNullOrEmpty(source.FontFamily) && !string.IsNullOrEmpty(dest.FontFamily))
            {
                if (!string.Equals(source.FontFamily, dest.FontFamily, StringComparison.OrdinalIgnoreCase))
                {
                    diffs.Add(new FormattingDifferenceItem
                    {
                        PropertyName = "Font Family",
                        SourceValue = source.FontFamily,
                        DestinationValue = dest.FontFamily
                    });
                }
            }

            // 2. Font Size
            if (source.FontSize > 0 && dest.FontSize > 0)
            {
                if (Math.Abs(source.FontSize - dest.FontSize) > 0.5f)
                {
                    diffs.Add(new FormattingDifferenceItem
                    {
                        PropertyName = "Font Size",
                        SourceValue = $"{source.FontSize:0.#} pt",
                        DestinationValue = $"{dest.FontSize:0.#} pt"
                    });
                }
            }

            // 3. Bold, Italic, Underline
            if (source.IsBold != dest.IsBold)
            {
                diffs.Add(new FormattingDifferenceItem
                {
                    PropertyName = "Bold",
                    SourceValue = source.IsBold ? "Bold" : "Regular",
                    DestinationValue = dest.IsBold ? "Bold" : "Regular"
                });
            }

            if (source.IsItalic != dest.IsItalic)
            {
                diffs.Add(new FormattingDifferenceItem
                {
                    PropertyName = "Italic",
                    SourceValue = source.IsItalic ? "Italic" : "Regular",
                    DestinationValue = dest.IsItalic ? "Italic" : "Regular"
                });
            }

            if (source.IsUnderline != dest.IsUnderline)
            {
                diffs.Add(new FormattingDifferenceItem
                {
                    PropertyName = "Underline",
                    SourceValue = source.IsUnderline ? "Underline" : "None",
                    DestinationValue = dest.IsUnderline ? "Underline" : "None"
                });
            }

            // 4. Colors
            if (!string.IsNullOrEmpty(source.FontColorHex) && !string.IsNullOrEmpty(dest.FontColorHex))
            {
                if (!string.Equals(source.FontColorHex, dest.FontColorHex, StringComparison.OrdinalIgnoreCase))
                {
                    diffs.Add(new FormattingDifferenceItem
                    {
                        PropertyName = "Font Color",
                        SourceValue = source.FontColorHex,
                        DestinationValue = dest.FontColorHex
                    });
                }
            }

            if (!string.IsNullOrEmpty(source.FillColorHex) && !string.IsNullOrEmpty(dest.FillColorHex))
            {
                if (!string.Equals(source.FillColorHex, dest.FillColorHex, StringComparison.OrdinalIgnoreCase))
                {
                    diffs.Add(new FormattingDifferenceItem
                    {
                        PropertyName = "Fill Color",
                        SourceValue = source.FillColorHex,
                        DestinationValue = dest.FillColorHex
                    });
                }
            }

            // 5. Alignment
            if (!string.IsNullOrEmpty(source.TextAlignment) && !string.IsNullOrEmpty(dest.TextAlignment))
            {
                if (!string.Equals(source.TextAlignment, dest.TextAlignment, StringComparison.OrdinalIgnoreCase))
                {
                    diffs.Add(new FormattingDifferenceItem
                    {
                        PropertyName = "Alignment",
                        SourceValue = source.TextAlignment,
                        DestinationValue = dest.TextAlignment
                    });
                }
            }

            // 6. Geometry: Width & Height
            if (Math.Abs(source.Width - dest.Width) > settings.SizeTolerancePoints ||
                Math.Abs(source.Height - dest.Height) > settings.SizeTolerancePoints)
            {
                diffs.Add(new FormattingDifferenceItem
                {
                    PropertyName = "Size",
                    SourceValue = $"{source.Width:0} x {source.Height:0} pt",
                    DestinationValue = $"{dest.Width:0} x {dest.Height:0} pt"
                });
            }

            // 7. Rotation
            if (Math.Abs(source.Rotation - dest.Rotation) > 2.0f)
            {
                diffs.Add(new FormattingDifferenceItem
                {
                    PropertyName = "Rotation",
                    SourceValue = $"{source.Rotation:0}°",
                    DestinationValue = $"{dest.Rotation:0}°"
                });
            }

            return diffs;
        }
    }
}

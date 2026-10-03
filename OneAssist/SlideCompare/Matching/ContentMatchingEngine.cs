using System;
using System.Security.Cryptography;
using System.Text;
using OneAssist.SlideCompare.Models;

namespace OneAssist.SlideCompare.Matching
{
    /// <summary>
    /// Generates content fingerprints and handles candidate filtering for high-performance comparison.
    /// </summary>
    public static class ContentMatchingEngine
    {
        /// <summary>
        /// Generates a stable SHA256 content fingerprint for a slide based on normalized title, body, tables, and shapes.
        /// </summary>
        public static string GenerateSlideFingerprint(SlideSnapshot slide)
        {
            if (slide == null) return string.Empty;

            var sb = new StringBuilder();
            sb.Append("T:").Append(slide.NormalizedTitle).Append("|");
            sb.Append("B:").Append(slide.NormalizedBodyText).Append("|");

            if (slide.Shapes != null)
            {
                foreach (var shp in slide.Shapes)
                {
                    sb.Append(shp.ShapeType).Append(":").Append(shp.NormalizedText).Append(";");
                }
            }

            using (var sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
                return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Generates a shape fingerprint based on type, normalized text, and approximate geometry.
        /// </summary>
        public static string GenerateShapeFingerprint(ShapeSnapshot shape)
        {
            if (shape == null) return string.Empty;

            int quantizedX = (int)(shape.Left / 10f);
            int quantizedY = (int)(shape.Top / 10f);
            string raw = $"{shape.ShapeType}:{shape.NormalizedText}:{quantizedX}:{quantizedY}";

            using (var sha = SHA256.Create())
            {
                byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(raw));
                return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant().Substring(0, 16);
            }
        }
    }
}

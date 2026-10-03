using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OneAssist.SlideCompare.Matching
{
    /// <summary>
    /// Utility for calculating similarity metrics between text, geometry, and structures.
    /// </summary>
    public static class SimilarityCalculator
    {
        private static readonly Regex WhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);
        private static readonly Regex PunctuationRegex = new Regex(@"[^\w\s]", RegexOptions.Compiled);

        /// <summary>
        /// Normalizes text by trimming, lowercasing, collapsing whitespace, and normalizing quotes/dashes.
        /// </summary>
        public static string NormalizeText(string text, bool stripPunctuation = false)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            string normalized = text.Trim().ToLowerInvariant();

            // Normalize quotes and dashes
            normalized = normalized
                .Replace('“', '"').Replace('”', '"').Replace('‘', '\'').Replace('’', '\'')
                .Replace('—', '-').Replace('–', '-').Replace('\r', ' ').Replace('\n', ' ');

            normalized = WhitespaceRegex.Replace(normalized, " ");

            if (stripPunctuation)
            {
                normalized = PunctuationRegex.Replace(normalized, "");
                normalized = WhitespaceRegex.Replace(normalized, " ");
            }

            return normalized.Trim();
        }

        /// <summary>
        /// Computes token-based Dice similarity between two text strings (0.0 to 1.0).
        /// </summary>
        public static double CalculateTextDiceSimilarity(string s1, string s2)
        {
            string n1 = NormalizeText(s1);
            string n2 = NormalizeText(s2);

            if (n1 == n2) return 1.0;
            if (string.IsNullOrEmpty(n1) || string.IsNullOrEmpty(n2)) return 0.0;

            var words1 = n1.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var words2 = n2.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (words1.Length == 0 && words2.Length == 0) return 1.0;
            if (words1.Length == 0 || words2.Length == 0) return 0.0;

            var set1 = new HashSet<string>(words1);
            var set2 = new HashSet<string>(words2);

            int intersection = set1.Count(w => set2.Contains(w));
            return (2.0 * intersection) / (set1.Count + set2.Count);
        }

        /// <summary>
        /// Computes normalized Levenshtein similarity between two strings (0.0 to 1.0).
        /// </summary>
        public static double CalculateLevenshteinSimilarity(string s1, string s2)
        {
            string n1 = NormalizeText(s1);
            string n2 = NormalizeText(s2);

            if (n1 == n2) return 1.0;
            if (string.IsNullOrEmpty(n1) && string.IsNullOrEmpty(n2)) return 1.0;
            if (string.IsNullOrEmpty(n1) || string.IsNullOrEmpty(n2)) return 0.0;

            int maxLen = Math.Max(n1.Length, n2.Length);
            if (maxLen == 0) return 1.0;

            // If strings are large, use word-level Dice similarity for speed
            if (maxLen > 250)
            {
                return CalculateTextDiceSimilarity(n1, n2);
            }

            int dist = ComputeLevenshteinDistance(n1, n2);
            return Math.Max(0.0, 1.0 - ((double)dist / maxLen));
        }

        private static int ComputeLevenshteinDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            var d = new int[n + 1, m + 1];

            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            return d[n, m];
        }

        /// <summary>
        /// Geometric position and size similarity (0.0 to 1.0).
        /// </summary>
        public static double CalculateGeometricSimilarity(
            float left1, float top1, float width1, float height1,
            float left2, float top2, float width2, float height2,
            float maxCanvasWidth = 960f, float maxCanvasHeight = 540f)
        {
            float dx = Math.Abs(left1 - left2);
            float dy = Math.Abs(top1 - top2);
            float dw = Math.Abs(width1 - width2);
            float dh = Math.Abs(height1 - height2);

            double posSim = Math.Max(0.0, 1.0 - ((dx / maxCanvasWidth * 0.5) + (dy / maxCanvasHeight * 0.5)));
            double sizeSim = Math.Max(0.0, 1.0 - ((dw / maxCanvasWidth * 0.5) + (dh / maxCanvasHeight * 0.5)));

            return (posSim * 0.6) + (sizeSim * 0.4);
        }
    }
}

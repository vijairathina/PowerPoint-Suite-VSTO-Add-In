using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using OneAssist.SlideCompare.Models;

namespace OneAssist.SlideCompare.Matching
{
    /// <summary>
    /// Word-level and character-level difference generator using LCS.
    /// </summary>
    public static class TextDiffEngine
    {
        private static readonly Regex WordSplitter = new Regex(@"(\s+|[^\w\s]+|\w+)", RegexOptions.Compiled);

        /// <summary>
        /// Generates a list of TextDiffTokens comparing source to destination text.
        /// </summary>
        public static List<TextDiffToken> ComputeDiff(string source, string destination, TextDiffMode mode = TextDiffMode.WordLevel)
        {
            source = source ?? string.Empty;
            destination = destination ?? string.Empty;

            if (source == destination)
            {
                if (string.IsNullOrEmpty(source))
                    return new List<TextDiffToken>();

                return new List<TextDiffToken>
                {
                    new TextDiffToken { DiffType = TextDiffType.Unchanged, Text = source }
                };
            }

            if (string.IsNullOrEmpty(source))
            {
                return new List<TextDiffToken>
                {
                    new TextDiffToken { DiffType = TextDiffType.Added, Text = destination }
                };
            }

            if (string.IsNullOrEmpty(destination))
            {
                return new List<TextDiffToken>
                {
                    new TextDiffToken { DiffType = TextDiffType.Deleted, Text = source }
                };
            }

            List<string> srcTokens = Tokenize(source, mode);
            List<string> dstTokens = Tokenize(destination, mode);

            // Compute LCS table
            int n = srcTokens.Count;
            int m = dstTokens.Count;
            int[,] lcs = new int[n + 1, m + 1];

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    if (string.Equals(srcTokens[i - 1], dstTokens[j - 1], StringComparison.OrdinalIgnoreCase))
                    {
                        lcs[i, j] = lcs[i - 1, j - 1] + 1;
                    }
                    else
                    {
                        lcs[i, j] = Math.Max(lcs[i - 1, j], lcs[i, j - 1]);
                    }
                }
            }

            // Backtrack to build raw operations
            var rawOps = new List<TextDiffToken>();
            int x = n;
            int y = m;

            while (x > 0 || y > 0)
            {
                if (x > 0 && y > 0 && string.Equals(srcTokens[x - 1], dstTokens[y - 1], StringComparison.OrdinalIgnoreCase))
                {
                    // Unchanged (favor destination casing)
                    rawOps.Add(new TextDiffToken { DiffType = TextDiffType.Unchanged, Text = dstTokens[y - 1] });
                    x--;
                    y--;
                }
                else if (y > 0 && (x == 0 || lcs[x, y - 1] >= lcs[x - 1, y]))
                {
                    // Added
                    rawOps.Add(new TextDiffToken { DiffType = TextDiffType.Added, Text = dstTokens[y - 1] });
                    y--;
                }
                else if (x > 0 && (y == 0 || lcs[x, y - 1] < lcs[x - 1, y]))
                {
                    // Deleted
                    rawOps.Add(new TextDiffToken { DiffType = TextDiffType.Deleted, Text = srcTokens[x - 1] });
                    x--;
                }
            }

            rawOps.Reverse();

            // Consolidate adjacent tokens of the same DiffType
            return ConsolidateTokens(rawOps);
        }

        private static List<string> Tokenize(string text, TextDiffMode mode)
        {
            var tokens = new List<string>();
            if (mode == TextDiffMode.CharacterLevel)
            {
                for (int i = 0; i < text.Length; i++)
                {
                    tokens.Add(text[i].ToString());
                }
                return tokens;
            }

            // Word-level tokenization preserving spaces and punctuation
            var matches = WordSplitter.Matches(text);
            foreach (Match m in matches)
            {
                tokens.Add(m.Value);
            }
            return tokens;
        }

        private static List<TextDiffToken> ConsolidateTokens(List<TextDiffToken> raw)
        {
            var result = new List<TextDiffToken>();
            if (raw.Count == 0) return result;

            var current = new TextDiffToken { DiffType = raw[0].DiffType, Text = raw[0].Text };

            for (int i = 1; i < raw.Count; i++)
            {
                if (raw[i].DiffType == current.DiffType)
                {
                    current.Text += raw[i].Text;
                }
                else
                {
                    result.Add(current);
                    current = new TextDiffToken { DiffType = raw[i].DiffType, Text = raw[i].Text };
                }
            }
            result.Add(current);
            return result;
        }

        /// <summary>
        /// Formats tokens into a human-readable visual difference string.
        /// Example: Revenue increased by [- 25%] [+ 30%] in 2025.
        /// </summary>
        public static string FormatInlineDiff(IEnumerable<TextDiffToken> tokens)
        {
            if (tokens == null) return string.Empty;

            var sb = new StringBuilder();
            foreach (var t in tokens)
            {
                switch (t.DiffType)
                {
                    case TextDiffType.Unchanged:
                        sb.Append(t.Text);
                        break;
                    case TextDiffType.Deleted:
                        sb.Append($"[- {t.Text.Trim()}] ");
                        break;
                    case TextDiffType.Added:
                        sb.Append($"[+ {t.Text.Trim()}] ");
                        break;
                    case TextDiffType.Modified:
                        sb.Append($"[~ {t.Text.Trim()}] ");
                        break;
                    case TextDiffType.Moved:
                        sb.Append($"[⇄ {t.Text.Trim()}] ");
                        break;
                }
            }
            return sb.ToString().Trim();
        }

        /// <summary>
        /// Produces a summary diff string showing only deleted and added lines.
        /// - 25%
        /// + 30%
        /// </summary>
        public static string FormatChangesOnly(IEnumerable<TextDiffToken> tokens)
        {
            if (tokens == null) return string.Empty;

            var deletedList = new List<string>();
            var addedList = new List<string>();

            foreach (var t in tokens)
            {
                string text = t.Text.Trim();
                if (string.IsNullOrEmpty(text)) continue;

                if (t.DiffType == TextDiffType.Deleted)
                {
                    deletedList.Add($"- {text}");
                }
                else if (t.DiffType == TextDiffType.Added)
                {
                    addedList.Add($"+ {text}");
                }
            }

            var sb = new StringBuilder();
            foreach (var d in deletedList) sb.AppendLine(d);
            foreach (var a in addedList) sb.AppendLine(a);

            return sb.ToString().Trim();
        }
    }
}

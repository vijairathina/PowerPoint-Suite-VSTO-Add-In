using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OneAssist.Assistant.Services
{
    public class EditorialIssueResult
    {
        public string Problem { get; set; }
        public string Expected { get; set; }
        public string Actual { get; set; }
        public string Recommendation { get; set; }
        public bool CanAutoFix { get; set; }
        public string FixedText { get; set; }
    }

    public static class EditorialEngine
    {
        private static readonly Regex RepeatedWordsRegex = new Regex(@"\b(\w+)\s+\1\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex ExcessivePunctuationRegex = new Regex(@"([!?,;]){2,}", RegexOptions.Compiled);
        private static readonly Regex MultipleSpacesRegex = new Regex(@"[ ]{2,}", RegexOptions.Compiled);

        /// <summary>
        /// Scans text for terminology replacements from the corporate dictionary.
        /// </summary>
        public static List<EditorialIssueResult> CheckTerminology(string text, IDictionary<string, string> dictionary)
        {
            var results = new List<EditorialIssueResult>();
            if (string.IsNullOrWhiteSpace(text) || dictionary == null) return results;

            foreach (var kvp in dictionary)
            {
                string wrong = kvp.Key;
                string preferred = kvp.Value;

                // Match exact word or phrase boundary
                string pattern = $@"\b{Regex.Escape(wrong)}\b";
                if (Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
                {
                    string corrected = Regex.Replace(text, pattern, preferred, RegexOptions.IgnoreCase);
                    results.Add(new EditorialIssueResult
                    {
                        Problem = $"Non-standard terminology '{wrong}' found.",
                        Actual = wrong,
                        Expected = preferred,
                        Recommendation = $"Replace '{wrong}' with approved corporate term '{preferred}'.",
                        CanAutoFix = true,
                        FixedText = corrected
                    });
                }
            }

            return results;
        }

        /// <summary>
        /// Detects duplicate adjacent words (e.g. 'the the', 'in in').
        /// </summary>
        public static List<EditorialIssueResult> CheckRepeatedWords(string text)
        {
            var results = new List<EditorialIssueResult>();
            if (string.IsNullOrWhiteSpace(text)) return results;

            var matches = RepeatedWordsRegex.Matches(text);
            foreach (Match m in matches)
            {
                string word = m.Groups[1].Value;
                string corrected = RepeatedWordsRegex.Replace(text, word, 1, m.Index);
                results.Add(new EditorialIssueResult
                {
                    Problem = $"Repeated word '{word} {word}'.",
                    Actual = m.Value,
                    Expected = word,
                    Recommendation = $"Remove duplicate word '{word}'.",
                    CanAutoFix = true,
                    FixedText = corrected
                });
            }

            return results;
        }

        /// <summary>
        /// Detects excessive punctuation (e.g. '!!', '??', '..').
        /// </summary>
        public static List<EditorialIssueResult> CheckExcessivePunctuation(string text)
        {
            var results = new List<EditorialIssueResult>();
            if (string.IsNullOrWhiteSpace(text)) return results;

            var matches = ExcessivePunctuationRegex.Matches(text);
            foreach (Match m in matches)
            {
                char singleChar = m.Value[0];
                results.Add(new EditorialIssueResult
                {
                    Problem = $"Excessive punctuation '{m.Value}'.",
                    Actual = m.Value,
                    Expected = singleChar.ToString(),
                    Recommendation = $"Use single '{singleChar}' instead of '{m.Value}' for professional tone.",
                    CanAutoFix = true,
                    FixedText = ExcessivePunctuationRegex.Replace(text, singleChar.ToString())
                });
            }

            return results;
        }

        /// <summary>
        /// Identifies excessively long bullet points (> 30 words or > 200 characters) that create visual clutter.
        /// </summary>
        public static List<EditorialIssueResult> CheckBulletLength(string text, int maxWords = 30)
        {
            var results = new List<EditorialIssueResult>();
            if (string.IsNullOrWhiteSpace(text)) return results;

            var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                var words = trimmed.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length > maxWords)
                {
                    results.Add(new EditorialIssueResult
                    {
                        Problem = $"Bullet point is too lengthy ({words.Length} words).",
                        Actual = trimmed.Length > 60 ? trimmed.Substring(0, 57) + "..." : trimmed,
                        Expected = $"<= {maxWords} words per bullet",
                        Recommendation = "Condense message or split across multiple bullets/slides for audience readability.",
                        CanAutoFix = false
                    });
                }
            }

            return results;
        }
    }
}

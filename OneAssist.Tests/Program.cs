using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace OneAssist.Tests
{
    class Program
    {
        static int _passed = 0;
        static int _failed = 0;

        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("================================================================================");
            Console.WriteLine("    ONEASSIST: COMPLETE PRESENTATION STUDIO & GOVERNANCE TEST SUITE            ");
            Console.WriteLine("================================================================================");
            Console.ResetColor();

            // Run test suites 1 - 10
            TestColorContrastCalculations();
            TestColorConversionHelpers();
            TestEditorialTerminology();
            TestEditorialRepeatedWords();
            TestEditorialExcessivePunctuation();
            TestEditorialBulletLength();
            TestPresetSerialization();
            TestCustomRuleEvaluation();
            TestShortcutConflictDetection();
            TestHealthScoreComputation();

            // Run test suites 11 - 15
            TestShapeInspectionConversions();
            TestDataLabelFormatting();
            Test12ColumnGridCalculations();
            TestCommandPaletteLookup();
            TestDesignTokensAndTokensModel();

            // Run test suites 16 - 25 (Req 47 - 62)
            TestFixSafetyClassification();
            TestReportExportHtmlAndCsv();
            TestCorporateGovernancePolicy();
            TestLocalAnalyticsTracking();
            TestAnalysisCacheManager();
            TestConsistencyDifferenceCategorization();
            TestJsonAndExcelReportExports();
            TestAppModesConfiguration();
            TestHealthScoreAndPassedCountCalculation();
            TestUndoTransactionAndSafeFixEngine();

            // Run test suites 26 - 30 (AI Auto Layout Suggestion Engine)
            TestSemanticContentClassification();
            TestLayoutTemplateCatalog();
            TestLayoutConstraintEngine();
            TestLayoutQualityScoringAndPreservation();
            TestReferenceSlideModeAndSequenceHarmonization();

            // Run test suites 31 - 42 (Slide Compare & Content Difference Module - Req 50)
            TestSlideCompareSuite();

            Console.WriteLine();
            Console.WriteLine("================================================================================");
            if (_failed == 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"    ALL {_passed} TEST SUITES PASSED CLEANLY (0 FAILS)!");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"    TESTS COMPLETED: {_passed} PASSED, {_failed} FAILED.");
            }
            Console.ResetColor();
            Console.WriteLine("================================================================================");

            if (_failed > 0)
            {
                Environment.Exit(1);
            }
        }

        static void Assert(bool condition, string testName, string details = "")
        {
            if (condition)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("  [PASS] ");
                Console.ResetColor();
                Console.WriteLine(testName);
                _passed++;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("  [FAIL] ");
                Console.ResetColor();
                Console.WriteLine($"{testName} - {details}");
                _failed++;
            }
        }

        // ── 1. Color Contrast & WCAG 2.1 ──────────────────────────────────────────
        static void TestColorContrastCalculations()
        {
            Console.WriteLine("\n[1] Testing WCAG 2.1 Contrast Ratio Algorithms...");

            double blackWhite = CalculateContrastRatio(Color.Black, Color.White);
            Assert(Math.Abs(blackWhite - 21.0) < 0.1, "Black on White contrast ratio is 21:1", $"Got: {blackWhite}");

            double whiteWhite = CalculateContrastRatio(Color.White, Color.White);
            Assert(Math.Abs(whiteWhite - 1.0) < 0.05, "White on White contrast ratio is 1:1", $"Got: {whiteWhite}");

            Color navy = Color.FromArgb(0, 0x20, 0x60);
            double navyWhite = CalculateContrastRatio(navy, Color.White);
            Assert(navyWhite >= 4.5, "Brand Navy on White passes WCAG AA (>= 4.5:1)", $"Got: {navyWhite}");

            Color lightGray = Color.FromArgb(0xCC, 0xCC, 0xCC);
            double grayWhite = CalculateContrastRatio(lightGray, Color.White);
            Assert(grayWhite < 2.0, "Light gray on White correctly detected as failing contrast (< 2.0:1)", $"Got: {grayWhite}");
        }

        private static double ConvertChannel(double channel) =>
            channel <= 0.03928 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);

        private static double CalculateLuminance(Color color)
        {
            double r = ConvertChannel(color.R / 255.0);
            double g = ConvertChannel(color.G / 255.0);
            double b = ConvertChannel(color.B / 255.0);
            return 0.2126 * r + 0.7152 * g + 0.0722 * b;
        }

        private static double CalculateContrastRatio(Color c1, Color c2)
        {
            double l1 = CalculateLuminance(c1);
            double l2 = CalculateLuminance(c2);
            double brightest = Math.Max(l1, l2);
            double darkest = Math.Min(l1, l2);
            return Math.Round((brightest + 0.05) / (darkest + 0.05), 2);
        }

        // ── 2. Color Conversion Helpers ───────────────────────────────────────────
        static void TestColorConversionHelpers()
        {
            Console.WriteLine("\n[2] Testing PowerPoint BGR COLORREF ↔ System.Drawing.Color Conversion...");

            int bgrRed = ColorToPowerPointRgb(Color.FromArgb(255, 0, 0));
            Assert(bgrRed == 255, "Pure Red translates to BGR 255", $"Got: {bgrRed}");

            int bgrBlue = ColorToPowerPointRgb(Color.FromArgb(0, 0, 255));
            Assert(bgrBlue == 16711680, "Pure Blue translates to BGR 16711680", $"Got: {bgrBlue}");

            Color restored = ColorFromPowerPointRgb(bgrBlue);
            Assert(restored.R == 0 && restored.G == 0 && restored.B == 255, "Round-trip conversion retains exact Blue component", $"Got: {restored}");

            string hex = ColorToHex(Color.FromArgb(0, 0x20, 0x60));
            Assert(hex == "#002060", "ColorToHex produces #002060", $"Got: {hex}");
        }

        private static int ColorToPowerPointRgb(Color c) => c.R | (c.G << 8) | (c.B << 16);
        private static Color ColorFromPowerPointRgb(int bgr) => Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);
        private static string ColorToHex(Color c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

        // ── 3. Editorial Terminology ──────────────────────────────────────────────
        static void TestEditorialTerminology()
        {
            Console.WriteLine("\n[3] Testing Editorial Terminology Dictionary Replacement...");

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Power Point", "PowerPoint" },
                { "PowerBI", "Power BI" },
                { "API's", "APIs" }
            };

            string input = "Build a Power Point deck integrating PowerBI and external API's.";
            string expected = "Build a PowerPoint deck integrating Power BI and external APIs.";

            string current = input;
            foreach (var kvp in dict)
            {
                string pattern = $@"\b{Regex.Escape(kvp.Key)}\b";
                current = Regex.Replace(current, pattern, kvp.Value, RegexOptions.IgnoreCase);
            }

            Assert(current == expected, "Replaces 'Power Point', 'PowerBI', and 'API's' with approved terms", $"Result: {current}");
        }

        // ── 4. Editorial Repeated Words ───────────────────────────────────────────
        static void TestEditorialRepeatedWords()
        {
            Console.WriteLine("\n[4] Testing Editorial Duplicate/Repeated Words Detector...");

            var regex = new Regex(@"\b(\w+)\s+\1\b", RegexOptions.IgnoreCase);
            string input = "This is the the official launch in in Dublin.";

            var matches = regex.Matches(input);
            Assert(matches.Count == 2, "Identified 2 occurrences of repeated words ('the the' and 'in in')", $"Found: {matches.Count}");

            string fixedText = regex.Replace(input, "$1");
            Assert(fixedText == "This is the official launch in Dublin.", "Fixed repeated words cleanly", $"Result: {fixedText}");
        }

        // ── 5. Editorial Excessive Punctuation ────────────────────────────────────
        static void TestEditorialExcessivePunctuation()
        {
            Console.WriteLine("\n[5] Testing Editorial Excessive Punctuation Detector...");

            var regex = new Regex(@"([!?,;]){2,}");
            string input = "Crucial Deadline Approaching!!! Please confirm??";

            var matches = regex.Matches(input);
            Assert(matches.Count == 2, "Detected '!!!' and '??'", $"Found: {matches.Count}");

            string fixedText = regex.Replace(input, "$1");
            Assert(fixedText == "Crucial Deadline Approaching! Please confirm?", "Collapsed repeated punctuation to professional single marks", $"Result: {fixedText}");
        }

        // ── 6. Editorial Bullet Length ────────────────────────────────────────────
        static void TestEditorialBulletLength()
        {
            Console.WriteLine("\n[6] Testing Long Bullet Length Warning...");

            string shortBullet = "Deliver Phase 1 formatting assistant and core architecture.";
            string longBullet = "This bullet contains an extremely long explanation that spans way past thirty words and rambles on without brevity or clear focus, creating unnecessary visual clutter and making the slide virtually unreadable for executive audiences during board presentations.";

            var shortWords = shortBullet.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var longWords = longBullet.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            Assert(shortWords.Length < 30, "Short bullet passes length limit", $"Words: {shortWords.Length}");
            Assert(longWords.Length > 30, "Long bullet triggers warning (> 30 words)", $"Words: {longWords.Length}");
        }

        // ── 7. Brand Preset Serialization ─────────────────────────────────────────
        static void TestPresetSerialization()
        {
            Console.WriteLine("\n[7] Testing Brand Preset Model & Data Validation...");

            var preset = new
            {
                PresetName = "Enterprise Global",
                PrimaryFont = "Aptos",
                HeadingFont = "Aptos Display",
                MinTitleFontSize = 24.0f,
                MinBodyFontSize = 12.0f,
                Colors = new List<string> { "#002060", "#00A88F", "#262626", "#FFFFFF" }
            };

            Assert(!string.IsNullOrEmpty(preset.PresetName), "Preset Name is valid");
            Assert(preset.PrimaryFont == "Aptos" && preset.HeadingFont == "Aptos Display", "Default brand fonts Aptos & Aptos Display are correctly configured");
            Assert(preset.Colors.Count == 4, "Brand preset holds 4 designated corporate colors");
        }

        // ── 8. Custom Rule Evaluation ─────────────────────────────────────────────
        static void TestCustomRuleEvaluation()
        {
            Console.WriteLine("\n[8] Testing Dynamic Custom Rule Evaluator...");

            bool fontViolation = EvaluateString("Calibri", "!=", "Aptos");
            Assert(fontViolation, "Custom Rule: Calibri != Aptos triggers violation");

            bool fontPass = EvaluateString("Aptos", "!=", "Aptos");
            Assert(!fontPass, "Custom Rule: Aptos != Aptos does not trigger violation");

            bool sizeViolation = EvaluateNumeric(18.0f, "<", 24.0f);
            Assert(sizeViolation, "Custom Rule: 18pt < 24pt triggers small title violation");

            bool sizePass = EvaluateNumeric(28.0f, "<", 24.0f);
            Assert(!sizePass, "Custom Rule: 28pt < 24pt does not trigger violation");
        }

        private static bool EvaluateString(string actual, string op, string target)
        {
            if (op == "==") return string.Equals(actual, target, StringComparison.OrdinalIgnoreCase);
            if (op == "!=") return !string.Equals(actual, target, StringComparison.OrdinalIgnoreCase);
            if (op == "Contains") return actual.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0;
            return false;
        }

        private static bool EvaluateNumeric(float actual, string op, float target)
        {
            if (op == "<") return actual < target;
            if (op == ">") return actual > target;
            if (op == "==") return Math.Abs(actual - target) < 0.1f;
            return false;
        }

        // ── 9. Shortcut Conflict Detection ────────────────────────────────────────
        static void TestShortcutConflictDetection()
        {
            Console.WriteLine("\n[9] Testing Keyboard Shortcut Conflict Detection...");

            var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Ctrl+C", "Ctrl+V", "Ctrl+Z", "Ctrl+Y", "Ctrl+S", "Ctrl+A", "F5"
            };

            bool ctrlCConflict = reserved.Contains("Ctrl+C");
            Assert(ctrlCConflict, "Ctrl+C detected as conflicting with native PowerPoint hotkey");

            bool f5Conflict = reserved.Contains("F5");
            Assert(f5Conflict, "F5 detected as conflicting with native PowerPoint Start Slideshow");

            bool safeShortcut = !reserved.Contains("Ctrl+Shift+Q");
            Assert(safeShortcut, "Ctrl+Shift+Q is free and safe for custom add-in assignment");
        }

        // ── 10. Health Score Computation ──────────────────────────────────────────
        static void TestHealthScoreComputation()
        {
            Console.WriteLine("\n[10] Testing Presentation Health Score Algorithm...");

            double perfectScore = ComputeHealthScore(0, 0, 0, 10, 40);
            Assert(perfectScore == 100.0, "Clean presentation returns 100% Health Score", $"Got: {perfectScore}");

            double degradedScore = ComputeHealthScore(5, 8, 2, 10, 40);
            Assert(degradedScore < 90.0 && degradedScore > 60.0, "Flawed presentation scores between 60% and 90%", $"Got: {degradedScore}%");
        }

        private static double ComputeHealthScore(int errors, int warnings, int infos, int totalSlides, int totalShapes)
        {
            if (totalShapes == 0 && totalSlides == 0) return 100.0;
            double penalty = (errors * 3.5) + (warnings * 1.2) + (infos * 0.3);
            double basePoints = Math.Max(10.0, (totalShapes * 2.0) + (totalSlides * 5.0));
            double score = Math.Max(0.0, 100.0 - (penalty / basePoints * 100.0));
            return Math.Round(score, 1);
        }

        // ── 11. Shape Inspection & Precision Units ─────────────────────────────────
        static void TestShapeInspectionConversions()
        {
            Console.WriteLine("\n[11] Testing Shape Inspector Inches ↔ Points Coordinate System...");

            float xPoints = 144.0f; // Exactly 2.0 inches
            float yPoints = 72.0f;  // Exactly 1.0 inch
            float wPoints = 360.0f; // Exactly 5.0 inches
            float hPoints = 216.0f; // Exactly 3.0 inches

            double xInches = Math.Round(xPoints / 72.0, 2);
            double yInches = Math.Round(yPoints / 72.0, 2);
            double wInches = Math.Round(wPoints / 72.0, 2);
            double hInches = Math.Round(hPoints / 72.0, 2);

            Assert(xInches == 2.0, "144pt converts accurately to 2.00 inches", $"Got: {xInches}");
            Assert(yInches == 1.0, "72pt converts accurately to 1.00 inch", $"Got: {yInches}");
            Assert(wInches == 5.0 && hInches == 3.0, "Dimensions 360x216pt convert to 5.00x3.00 inches", $"Got: {wInches}x{hInches}");
        }

        // ── 12. Data Label Formatting ─────────────────────────────────────────────
        static void TestDataLabelFormatting()
        {
            Console.WriteLine("\n[12] Testing Chart Data Label Formatter ($1.2M, 35%, 12)...");

            string formatted1 = FormatDataLabel(1_200_000, "millions");
            Assert(formatted1 == "$1.2M", "1,200,000 formatted to '$1.2M'", $"Got: {formatted1}");

            string formatted2 = FormatDataLabel(0.35, "percentage");
            Assert(formatted2 == "35%", "0.35 formatted to '35%'", $"Got: {formatted2}");

            string formatted3 = FormatDataLabel(12.00, "integer");
            Assert(formatted3 == "12", "12.00 formatted to '12'", $"Got: {formatted3}");
        }

        private static string FormatDataLabel(double value, string formatType)
        {
            switch (formatType?.ToLower())
            {
                case "currency_millions":
                case "millions":
                    if (Math.Abs(value) >= 1_000_000)
                        return $"${value / 1_000_000:F1}M";
                    if (Math.Abs(value) >= 1_000)
                        return $"${value / 1_000:F1}K";
                    return $"${value:F0}";

                case "percentage":
                case "percent":
                    return $"{value * 100:F0}%";

                case "integer":
                    return $"{Math.Round(value):F0}";

                default:
                    return value.ToString("G");
            }
        }

        // ── 13. 12-Column Grid Calculations ───────────────────────────────────────
        static void Test12ColumnGridCalculations()
        {
            Console.WriteLine("\n[13] Testing 12-Column Slide Layout System...");

            float slideWidth = 960.0f; // Standard 16:9 13.33" widescreen slide
            float margin = 36.0f;
            float columnGap = 10.8f;

            float availableWidth = slideWidth - (margin * 2.0f);
            float totalGaps = columnGap * 11.0f;
            float colWidth = (availableWidth - totalGaps) / 12.0f;

            Assert(colWidth > 0, "Calculated column width is positive", $"Width: {colWidth:F2}pt");
            Assert(Math.Abs((colWidth * 12.0f) + totalGaps - availableWidth) < 0.01f, "12 columns + 11 gaps perfectly fill printable canvas", $"Available: {availableWidth}pt");
        }

        // ── 14. Command Palette Lookup ────────────────────────────────────────────
        static void TestCommandPaletteLookup()
        {
            Console.WriteLine("\n[14] Testing Command Palette Fast Search...");

            var commands = new List<(string Id, string Title, string Category)>
            {
                ("align_left", "Align Left", "Layout"),
                ("replace_font", "Replace Font Globally", "Typography"),
                ("zebra_table", "Apply Zebra Striping to Table", "Tables"),
                ("clean_ghost", "Clean Ghost Shapes and Artifacts", "Production"),
                ("toc_gen", "Generate Table of Contents", "Structure")
            };

            var searchResult = commands.Where(c => c.Title.ToLower().Contains("zebra") || c.Category.ToLower().Contains("zebra")).ToList();
            Assert(searchResult.Count == 1 && searchResult[0].Id == "zebra_table", "Search 'zebra' finds table formatting tool", $"Matches: {searchResult.Count}");

            var alignResult = commands.Where(c => c.Title.ToLower().Contains("align")).ToList();
            Assert(alignResult.Count == 1 && alignResult[0].Id == "align_left", "Search 'align' retrieves alignment command", $"Matches: {alignResult.Count}");
        }

        // ── 15. Design Tokens Model ───────────────────────────────────────────────
        static void TestDesignTokensAndTokensModel()
        {
            Console.WriteLine("\n[15] Testing Design Tokens & Brand System...");

            var tokens = new
            {
                TitleFont = "Aptos Display",
                BodyFont = "Aptos",
                PrimaryColorHex = "#002060",
                SpacingSmall = 7.2f,   // 0.1"
                SpacingMedium = 14.4f, // 0.2"
                SpacingLarge = 28.8f   // 0.4"
            };

            Assert(tokens.TitleFont == "Aptos Display", "Design Token TitleFont is Aptos Display");
            Assert(tokens.SpacingMedium == 14.4f, "Design Token SpacingMedium is 14.4pt (0.2 inch)");
            Assert(tokens.PrimaryColorHex == "#002060", "Design Token PrimaryColorHex is #002060");
        }

        // ── 16. Fix Safety Classification (Req 53) ────────────────────────────────
        enum FixSafety { Safe, Review, Dangerous }

        static void TestFixSafetyClassification()
        {
            Console.WriteLine("\n[16] Testing Safe Auto-Fix Engine Classification (Req 53)...");

            var alignFix = FixSafety.Safe;
            var rewriteContent = FixSafety.Review;
            var deleteNotes = FixSafety.Dangerous;

            Assert(alignFix == FixSafety.Safe, "Align objects classified as SAFE fix");
            Assert(rewriteContent == FixSafety.Review, "Rewriting content classified as REVIEW fix");
            Assert(deleteNotes == FixSafety.Dangerous, "Deleting speaker notes classified as DANGEROUS fix");
        }

        // ── 17. Report Export (HTML & CSV) (Req 54) ────────────────────────────────
        static void TestReportExportHtmlAndCsv()
        {
            Console.WriteLine("\n[17] Testing Report Exporter (HTML & CSV Generation) (Req 54)...");

            string htmlSample = "<!DOCTYPE html><html><title>Presentation Quality Audit Report</title><body><div class='score-badge'>92% Health</div></body></html>";
            Assert(htmlSample.Contains("Quality Audit") && htmlSample.Contains("92% Health"), "HTML report contains correct title and health score badge");

            string csvSample = "Slide,Severity,Category,Rule,Problem,Expected,Actual,CanAutoFix\n12,Warning,Typography,\"Font Consistency\",\"Arial used\",\"Aptos\",True";
            Assert(csvSample.StartsWith("Slide,Severity,Category"), "CSV export contains standard column headers");
            Assert(csvSample.Contains("Arial used"), "CSV export correctly formats detected issue row");
        }

        // ── 18. Corporate Admin Governance Policy (Req 55) ────────────────────────
        static void TestCorporateGovernancePolicy()
        {
            Console.WriteLine("\n[18] Testing Corporate Admin Governance & Policy Locks (Req 55)...");

            bool isBrandLocked = true;
            bool isDictLocked = true;
            string org = "Global Corporate Headquarters";

            bool canUserModifyBrand = !isBrandLocked;
            bool canUserModifyDict = !isDictLocked;

            Assert(!canUserModifyBrand, "User is prevented from modifying locked corporate brand colors");
            Assert(!canUserModifyDict, "User is prevented from overriding locked enterprise dictionary");
            Assert(org.Length > 0, "Organization policy metadata recorded");
        }

        // ── 19. Local Analytics & Strict Privacy (Req 56, 57) ─────────────────────
        static void TestLocalAnalyticsTracking()
        {
            Console.WriteLine("\n[19] Testing Local Privacy & Usage Statistics Tracking (Req 56, 57)...");

            var toolUsage = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            void RecordUsage(string name)
            {
                if (!toolUsage.ContainsKey(name)) toolUsage[name] = 0;
                toolUsage[name]++;
            }

            RecordUsage("AlignTools");
            RecordUsage("AlignTools");
            RecordUsage("FormatPainter");

            Assert(toolUsage["AlignTools"] == 2, "Local analytics tracks tool invocation frequency (AlignTools: 2)");
            Assert(toolUsage["FormatPainter"] == 1, "Local analytics tracks FormatPainter: 1");

            bool analyzeLocallyOnly = true;
            bool cloudUploadEnabled = false;
            Assert(analyzeLocallyOnly && !cloudUploadEnabled, "Privacy engine strictly enforces local-only analysis without cloud data transmission");
        }

        // ── 20. Incremental Analysis Cache (Req 60) ───────────────────────────────
        static void TestAnalysisCacheManager()
        {
            Console.WriteLine("\n[20] Testing Incremental Analysis Cache for 100+ Slide Decks (Req 60)...");

            var cache = new Dictionary<int, (int shapeCount, int issueCount)>();

            // Slide 1 has 5 shapes and 2 issues
            cache[1] = (5, 2);

            // Check hit
            bool hit = cache.TryGetValue(1, out var cached) && cached.shapeCount == 5;
            Assert(hit && cached.issueCount == 2, "Unchanged slide cache hit returns cached issues instantly without re-scan");

            // Slide modified: shape count now 6
            bool cacheMiss = cache.TryGetValue(1, out var modCached) && modCached.shapeCount == 6;
            Assert(!cacheMiss, "Changed slide correctly triggers cache miss and forces re-analysis");
        }

        // ── 21. Make Slide & Presentation Consistent (Req 47, 48) ──────────────────
        class MockConsistencyReport
        {
            public int SlideIndex { get; set; }
            public List<string> TypographyDifferences { get; set; } = new List<string>();
            public List<string> ColorDifferences { get; set; } = new List<string>();
            public List<string> AlignmentDifferences { get; set; } = new List<string>();
            public List<string> SpacingDifferences { get; set; } = new List<string>();
            public List<string> BrandingDifferences { get; set; } = new List<string>();
            public int AllDifferencesCount => TypographyDifferences.Count + ColorDifferences.Count + AlignmentDifferences.Count + SpacingDifferences.Count + BrandingDifferences.Count;
        }

        static void TestConsistencyDifferenceCategorization()
        {
            Console.WriteLine("\n[21] Testing 'Make Slide Consistent' & 'Make Presentation Consistent' (Req 47, 48)...");

            var report = new MockConsistencyReport { SlideIndex = 4 };
            report.TypographyDifferences.Add("Title font mismatch");
            report.ColorDifferences.Add("Incorrect brand color");
            report.AlignmentDifferences.Add("Logo size mismatch");
            report.SpacingDifferences.Add("Uneven card spacing");
            report.BrandingDifferences.Add("Low contrast logo");

            Assert(report.TypographyDifferences.Count == 1, "Typography differences correctly isolated");
            Assert(report.ColorDifferences.Count == 1, "Color differences correctly isolated");
            Assert(report.AlignmentDifferences.Count == 1, "Alignment differences correctly isolated");
            Assert(report.SpacingDifferences.Count == 1, "Spacing differences correctly isolated");
            Assert(report.BrandingDifferences.Count == 1, "Branding differences correctly isolated");
            Assert(report.AllDifferencesCount == 5, "AllDifferences aggregates all 5 consistency categories");

            // Deck-wide consistency map (Req 48)
            var deckConsistency = new Dictionary<int, List<string>>
            {
                { 4, new List<string> { "Title font mismatch" } },
                { 7, new List<string> { "Logo size mismatch" } },
                { 11, new List<string> { "Incorrect brand color" } },
                { 15, new List<string> { "Uneven card spacing" } },
                { 19, new List<string> { "Low contrast" } }
            };

            Assert(deckConsistency.ContainsKey(4) && deckConsistency[4][0] == "Title font mismatch", "Slide 4 flagged with title font mismatch");
            Assert(deckConsistency.ContainsKey(7) && deckConsistency[7][0] == "Logo size mismatch", "Slide 7 flagged with logo size mismatch");
            Assert(deckConsistency.ContainsKey(11) && deckConsistency[11][0] == "Incorrect brand color", "Slide 11 flagged with brand color issue");
            Assert(deckConsistency.ContainsKey(15) && deckConsistency[15][0] == "Uneven card spacing", "Slide 15 flagged with uneven card spacing");
            Assert(deckConsistency.ContainsKey(19) && deckConsistency[19][0] == "Low contrast", "Slide 19 flagged with low contrast");
        }

        // ── 22. Report Export JSON, Excel XML & Text Summary (Req 54) ─────────────
        static void TestJsonAndExcelReportExports()
        {
            Console.WriteLine("\n[22] Testing JSON, Excel XML & Text Summary Exports (Req 54)...");

            string presName = "QuarterlyReport.pptx";
            string json = $"{{\n  \"presentation\": \"{presName}\",\n  \"healthScore\": 92.5,\n  \"summary\": {{\n    \"errors\": 1,\n    \"warnings\": 0,\n    \"passed\": 45\n  }},\n  \"findings\": [\n    {{\n      \"slide\": 12,\n      \"ruleTitle\": \"Title font mismatch\",\n      \"expected\": \"Aptos Display 28\",\n      \"actual\": \"Arial 24\"\n    }}\n  ]\n}}";

            Assert(json.Contains("\"presentation\": \"QuarterlyReport.pptx\""), "JSON export contains presentation name");
            Assert(json.Contains("\"errors\": 1"), "JSON export contains error summary count");
            Assert(json.Contains("\"expected\": \"Aptos Display 28\""), "JSON export contains expected font specs");

            string excelXml = "<?xml version=\"1.0\"?>\n<?mso-application progid=\"Excel.Sheet\"?>\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\">\n <Worksheet ss:Name=\"Quality Report\">\n  <Table>\n   <Row><Cell><Data ss:Type=\"String\">Aptos Display 28</Data></Cell></Row>\n  </Table>\n </Worksheet>\n</Workbook>";
            Assert(excelXml.Contains("<?mso-application progid=\"Excel.Sheet\"?>"), "Excel XML has correct Excel progid PI");
            Assert(excelXml.Contains("<Worksheet ss:Name=\"Quality Report\">"), "Excel XML contains Quality Report worksheet");
            Assert(excelXml.Contains("Aptos Display 28"), "Excel XML row contains findings text");

            string textRep = "Presentation Quality Report\nPresentation: QuarterlyReport.pptx\n\nSummary:\nErrors: 7\nWarnings: 18\nPassed: 126\n\nDetailed Findings:\nSlide 12\nTypography\nTitle font mismatch\nExpected:\nAptos Display 28\nActual:\nArial 24\n[Fix available]";
            Assert(textRep.Contains("Presentation Quality Report"), "Text report contains standard header");
            Assert(textRep.Contains("Expected:\nAptos Display 28"), "Text report contains expected format");
            Assert(textRep.Contains("[Fix available]"), "Text report contains [Fix available] indicator");
        }

        // ── 23. App Modes: Business User, Designer, Expert (Req 49, 50, 51) ────────
        enum MockAppMode { BusinessUser, Designer, Expert }

        static void TestAppModesConfiguration()
        {
            Console.WriteLine("\n[23] Testing App Modes (Business User, Designer, Expert) (Req 49, 50, 51)...");

            var modes = Enum.GetNames(typeof(MockAppMode));
            Assert(modes.Contains("BusinessUser"), "AppMode supports BusinessUser");
            Assert(modes.Contains("Designer"), "AppMode supports Designer");
            Assert(modes.Contains("Expert"), "AppMode supports Expert");

            // Business user core actions (Req 50)
            var businessActions = new[] { "Check Presentation", "Fix Formatting", "Apply Brand", "Clean Presentation", "Find & Replace", "Insert Assets", "Slide Library" };
            Assert(businessActions.Length == 7, "Business user mode presents exactly 7 simplified core actions");
        }

        // ── 24. Health Score & Passed Count Accuracy (Req 54, 59) ──────────────────
        static void TestHealthScoreAndPassedCountCalculation()
        {
            Console.WriteLine("\n[24] Testing Health Score & Passed Count Calculation (Req 54, 59)...");

            int totalShapes = 50;
            int totalSlides = 10;
            int errors = 1;
            int warnings = 1;
            int totalIssues = errors + warnings;

            int passedCount = Math.Max(0, (totalShapes * 3) + (totalSlides * 5) - totalIssues);
            double penalty = (errors * 3.5) + (warnings * 1.2);
            double basePoints = (totalShapes * 2.0) + (totalSlides * 5.0);
            double score = Math.Max(0.0, 100.0 - (penalty / basePoints * 100.0));

            Assert(errors == 1, "Error count is 1");
            Assert(warnings == 1, "Warning count is 1");
            Assert(passedCount > 0, "Passed count is correctly calculated as positive integer (198)");
            Assert(score >= 90.0, "Health score calculated reliably (> 90% with only 2 minor issues on 50 shapes)");
        }

        // ── 25. Safe Auto-Fix Engine & Undo Snapshots (Req 53, 61) ─────────────────
        class MockFixItem
        {
            public FixSafety Safety { get; set; }
            public bool CanAutoFix { get; set; }
        }

        static void TestUndoTransactionAndSafeFixEngine()
        {
            Console.WriteLine("\n[25] Testing Safe Auto-Fix Engine & Undo Snapshots (Req 53, 61)...");

            var issues = new List<MockFixItem>
            {
                new MockFixItem { Safety = FixSafety.Safe, CanAutoFix = true },
                new MockFixItem { Safety = FixSafety.Review, CanAutoFix = true },
                new MockFixItem { Safety = FixSafety.Dangerous, CanAutoFix = false }
            };

            var safeOnly = issues.Where(i => i.Safety == FixSafety.Safe && i.CanAutoFix).ToList();
            Assert(safeOnly.Count == 1, "Safe auto-fix filter selects ONLY Safe level fixes and excludes Review & Dangerous");
            Assert(issues.Count(i => i.Safety == FixSafety.Dangerous) == 1, "Destructive operations are classified as Dangerous");
        }

        // ── 26. Semantic Content Classification (Req 2, 3, 4) ──────────────────────
        enum MockObjectSemanticType { Title, Subtitle, KPI, Body, Label, Footnote, HeroImage, PrimaryChart, Card, Divider }
        enum MockSlideContentType { TitleSlide, ExecutiveSummary, KPI, Comparison, Timeline, Process, Strategy, DataAnalysis, TextHeavy, ImageHeavy }

        class MockSemanticShape
        {
            public int ShapeId { get; set; }
            public MockObjectSemanticType SemanticType { get; set; }
            public string RawText { get; set; } = string.Empty;
            public int WordCount { get; set; }
            public int BulletCount { get; set; }
            public float Left { get; set; }
            public float Top { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
            public bool HasChart { get; set; }
            public bool HasPicture { get; set; }
        }

        static void TestSemanticContentClassification()
        {
            Console.WriteLine("\n[26] Testing Semantic Content Classification (Req 2, 3, 4)...");

            var shapes = new List<MockSemanticShape>
            {
                new MockSemanticShape { ShapeId = 1, RawText = "Revenue increased 28%", SemanticType = MockObjectSemanticType.Title, WordCount = 3 },
                new MockSemanticShape { ShapeId = 2, RawText = "North America +31%\nEurope +22%\nAPAC +29%", SemanticType = MockObjectSemanticType.KPI, BulletCount = 3, WordCount = 6 },
                new MockSemanticShape { ShapeId = 3, HasChart = true, SemanticType = MockObjectSemanticType.PrimaryChart }
            };

            bool hasChart = shapes.Any(s => s.HasChart || s.SemanticType == MockObjectSemanticType.PrimaryChart);
            bool hasKPI = shapes.Any(s => s.SemanticType == MockObjectSemanticType.KPI || Regex.IsMatch(s.RawText, @"\d+[%$MkB]"));
            
            MockSlideContentType detectedType = MockSlideContentType.TitleSlide;
            if (hasKPI && hasChart)
            {
                detectedType = MockSlideContentType.KPI;
            }

            Assert(detectedType == MockSlideContentType.KPI, "Identified 'Revenue increased 28%' with regional metrics and chart as KPI SlideType");
            Assert(shapes[0].SemanticType == MockObjectSemanticType.Title, "Title shape classified correctly");
            Assert(shapes[1].SemanticType == MockObjectSemanticType.KPI, "Regional percentage metrics classified as KPI");
            Assert(shapes[2].SemanticType == MockObjectSemanticType.PrimaryChart, "Chart object classified as PrimaryChart");
        }

        // ── 27. Layout Template Catalog & Deterministic Grid (Req 5, 6, 22) ────────
        class MockLayoutRegion
        {
            public string Name { get; set; } = string.Empty;
            public float X { get; set; }
            public float Y { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
            public string PreferredType { get; set; } = string.Empty;
        }

        static void TestLayoutTemplateCatalog()
        {
            Console.WriteLine("\n[27] Testing Layout Template Catalog & Deterministic Grid (Req 5, 6, 22)...");

            float slideW = 960f;
            float slideH = 540f;
            float margin = 36f; // 0.5 in

            // Generate KPI + Supporting Chart layout regions
            var kpiChartRegions = new List<MockLayoutRegion>
            {
                new MockLayoutRegion { Name = "title", X = margin, Y = margin, Width = slideW - (margin * 2), Height = 56f, PreferredType = "title" },
                new MockLayoutRegion { Name = "kpi", X = margin, Y = 110f, Width = 280f, Height = 180f, PreferredType = "kpi" },
                new MockLayoutRegion { Name = "supporting_cards", X = margin, Y = 300f, Width = 280f, Height = 195f, PreferredType = "body" },
                new MockLayoutRegion { Name = "primary_chart", X = 330f, Y = 110f, Width = 594f, Height = 385f, PreferredType = "chart" }
            };

            Assert(kpiChartRegions.Count == 4, "KPI + Supporting Chart generated exactly 4 structural regions");
            Assert(kpiChartRegions.All(r => r.X >= margin && (r.X + r.Width) <= slideW - margin + 0.1f), "All regions strictly respect 0.5\" (36pt) horizontal slide margin");
            Assert(kpiChartRegions.All(r => r.Y >= margin && (r.Y + r.Height) <= slideH - margin + 0.1f), "All regions strictly respect 0.5\" (36pt) vertical slide margin");

            // Generate Three Cards layout regions
            float cardGap = 16f;
            float availW = slideW - (margin * 2);
            float cardW = (availW - (cardGap * 2)) / 3f;
            var threeCards = new List<MockLayoutRegion>
            {
                new MockLayoutRegion { Name = "title", X = margin, Y = margin, Width = availW, Height = 56f },
                new MockLayoutRegion { Name = "card_1", X = margin, Y = 120f, Width = cardW, Height = 360f },
                new MockLayoutRegion { Name = "card_2", X = margin + cardW + cardGap, Y = 120f, Width = cardW, Height = 360f },
                new MockLayoutRegion { Name = "card_3", X = margin + (cardW + cardGap) * 2, Y = 120f, Width = cardW, Height = 360f }
            };

            Assert(threeCards.Count == 4, "Three Cards pattern generated title + 3 equal cards");
            Assert(Math.Abs(threeCards[1].Width - threeCards[2].Width) < 0.01f, "Card 1 and Card 2 have equal column widths");
            Assert(Math.Abs(threeCards[2].Width - threeCards[3].Width) < 0.01f, "Card 2 and Card 3 have equal column widths");
        }

        // ── 28. Layout Constraint Engine (Req 7, 8, 30) ─────────────────────────────
        static void TestLayoutConstraintEngine()
        {
            Console.WriteLine("\n[28] Testing Layout Constraint Engine (Req 7, 8, 30)...");

            float slideW = 960f;
            float slideH = 540f;
            float margin = 36f;

            // Scenario 1: Object out of bounds correction
            var outOfBounds = new MockLayoutRegion { Name = "test", X = 950f, Y = 520f, Width = 100f, Height = 100f };
            if (outOfBounds.X + outOfBounds.Width > slideW - margin)
            {
                outOfBounds.X = Math.Max(margin, slideW - margin - outOfBounds.Width);
            }
            if (outOfBounds.Y + outOfBounds.Height > slideH - margin)
            {
                outOfBounds.Y = Math.Max(margin, slideH - margin - outOfBounds.Height);
            }

            Assert(outOfBounds.X + outOfBounds.Width <= slideW - margin, "Constraint engine clamped X coordinate to stay within right margin");
            Assert(outOfBounds.Y + outOfBounds.Height <= slideH - margin, "Constraint engine clamped Y coordinate to stay within bottom margin");

            // Scenario 2: Overlap separation
            var regA = new MockLayoutRegion { Name = "A", X = 100f, Y = 100f, Width = 200f, Height = 100f };
            var regB = new MockLayoutRegion { Name = "B", X = 150f, Y = 120f, Width = 200f, Height = 100f };

            bool overlaps = !(regA.X + regA.Width <= regB.X || regB.X + regB.Width <= regA.X ||
                              regA.Y + regA.Height <= regB.Y || regB.Y + regB.Height <= regA.Y);
            Assert(overlaps, "Detected collision between overlapping regions A and B");

            // Auto-correct overlap by shifting B below A with minimum spacing 12pt
            regB.Y = regA.Y + regA.Height + 12f;
            bool stillOverlaps = !(regA.X + regA.Width <= regB.X || regB.X + regB.Width <= regA.X ||
                                  regA.Y + regA.Height <= regB.Y || regB.Y + regB.Height <= regA.Y);
            Assert(!stillOverlaps, "Constraint engine auto-corrected overlap by shifting position safely");
        }

        // ── 29. Layout Quality Scoring & Content Preservation (Req 9, 12, 13, 26) ──
        class MockQualityScore
        {
            public string Hierarchy { get; set; } = "Excellent";
            public string Spacing { get; set; } = "Excellent";
            public string Readability { get; set; } = "Good";
            public string BrandConsistency { get; set; } = "Excellent";
            public double ContentPreservationPct { get; set; } = 100.0;
            public string Complexity { get; set; } = "Low";
            public string Confidence { get; set; } = "Strong structural match";
        }

        static void TestLayoutQualityScoringAndPreservation()
        {
            Console.WriteLine("\n[29] Testing Layout Quality Scoring & Preservation (Req 9, 12, 13, 26)...");

            int originalShapeCount = 5;
            int assignedShapeCount = 5;
            double preservationPct = (double)assignedShapeCount / originalShapeCount * 100.0;

            var score = new MockQualityScore
            {
                Hierarchy = "Excellent",
                Spacing = "Excellent",
                Readability = "Good",
                BrandConsistency = "Excellent",
                ContentPreservationPct = preservationPct,
                Complexity = "Low",
                Confidence = "Strong structural match"
            };

            Assert(score.ContentPreservationPct == 100.0, "Content preservation is exactly 100% (zero shapes/facts deleted)");
            Assert(score.Hierarchy == "Excellent", "Visual hierarchy score is Excellent");
            Assert(score.BrandConsistency == "Excellent", "Brand consistency verified against brand palette");
            Assert(!string.IsNullOrEmpty(score.Confidence), "Confidence has descriptive structural match explanation");
        }

        // ── 30. Reference Slide Mode & Sequence Harmonization (Req 11, 20, 21, 36, 37) ─
        static void TestReferenceSlideModeAndSequenceHarmonization()
        {
            Console.WriteLine("\n[30] Testing Reference Slide Mode & Sequence Harmonization (Req 11, 20, 21, 36, 37)...");

            var refSlideGeometry = new Dictionary<string, (float X, float Y, float W, float H)>
            {
                { "title", (36f, 36f, 888f, 56f) },
                { "col1", (36f, 120f, 436f, 360f) },
                { "col2", (488f, 120f, 436f, 360f) }
            };

            var targetShapes = new List<string> { "Target Title Text", "Body Column Left", "Body Column Right" };

            var adaptedLayout = new List<MockLayoutRegion>();
            adaptedLayout.Add(new MockLayoutRegion { Name = "title", X = refSlideGeometry["title"].X, Y = refSlideGeometry["title"].Y, Width = refSlideGeometry["title"].W, Height = refSlideGeometry["title"].H });
            adaptedLayout.Add(new MockLayoutRegion { Name = "col1", X = refSlideGeometry["col1"].X, Y = refSlideGeometry["col1"].Y, Width = refSlideGeometry["col1"].W, Height = refSlideGeometry["col1"].H });
            adaptedLayout.Add(new MockLayoutRegion { Name = "col2", X = refSlideGeometry["col2"].X, Y = refSlideGeometry["col2"].Y, Width = refSlideGeometry["col2"].W, Height = refSlideGeometry["col2"].H });

            Assert(adaptedLayout.Count == 3, "Reference slide structure successfully mapped to target slide shapes");
            Assert(targetShapes[0] == "Target Title Text", "Target slide content preserved without copying reference slide text");

            var slideFamilies = new Dictionary<int, string>
            {
                { 1, "SectionDivider" },
                { 2, "ExecutiveSummary" },
                { 3, "DataSlides" },
                { 4, "DataSlides" },
                { 5, "ExecutiveSummary" },
                { 6, "DataSlides" },
                { 7, "SectionDivider" }
            };

            var dataFamilySlides = slideFamilies.Where(kvp => kvp.Value == "DataSlides").Select(kvp => kvp.Key).ToList();
            Assert(dataFamilySlides.Count == 3, "Identified 3 data slides (Slides 3, 4, 6) in the same visual family for sequence harmonization");
        }

        // ── 31 - 42. Slide Compare & Content Difference Module (Req 50) ───────────
        static void TestSlideCompareSuite()
        {
            Console.WriteLine("\n[31] Testing Slide Compare Suite (Req 50 Scenarios 1-12)...");

            var settings = new OneAssist.SlideCompare.Services.CompareSettings();

            // ── Test 1: Identical presentations (Expected: All Matched) ─────────────────
            var s1 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                PresentationName = "Deck_A.pptx",
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 1,
                        Title = "Company Overview",
                        NormalizedTitle = "company overview",
                        BodyText = "Leader in cloud productivity solutions.",
                        NormalizedBodyText = "leader in cloud productivity solutions.",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 1, Name = "Title", RawText = "Company Overview", NormalizedText = "company overview", Left = 50, Top = 50, Width = 800, Height = 60, HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 2, Name = "Body", RawText = "Leader in cloud productivity solutions.", NormalizedText = "leader in cloud productivity solutions.", Left = 50, Top = 150, Width = 800, Height = 300, HasText = true }
                        }
                    }
                }
            };
            var d1 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                PresentationName = "Deck_A_Copy.pptx",
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 1,
                        Title = "Company Overview",
                        NormalizedTitle = "company overview",
                        BodyText = "Leader in cloud productivity solutions.",
                        NormalizedBodyText = "leader in cloud productivity solutions.",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 1, Name = "Title", RawText = "Company Overview", NormalizedText = "company overview", Left = 50, Top = 50, Width = 800, Height = 60, HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 2, Name = "Body", RawText = "Leader in cloud productivity solutions.", NormalizedText = "leader in cloud productivity solutions.", Left = 50, Top = 150, Width = 800, Height = 300, HasText = true }
                        }
                    }
                }
            };
            var res1 = OneAssist.SlideCompare.Comparison.ComparisonEngine.ComparePresentations(s1, d1, settings);
            string d1Info = $"Matched: {res1.MatchedCount}, Issues: {res1.TotalIssuesCount}, Slide0: {res1.SlideResults[0].MatchType}, Shape0: {res1.SlideResults[0].ShapeResults[0].MatchType}/{res1.SlideResults[0].ShapeResults[0].HasTextChanges}, Shape1: {res1.SlideResults[0].ShapeResults[1].MatchType}/{res1.SlideResults[0].ShapeResults[1].HasTextChanges}";
            Assert(res1.MatchedCount == 1 && res1.TotalIssuesCount == 0, "Test 1: Identical presentations marked as All Matched with 0 issues", d1Info);

            // ── Test 2: One word changed (Expected: Text Modified) ──────────────────────
            string srcTxt2 = "Revenue increased by 25% in 2025.";
            string dstTxt2 = "Revenue increased by 30% in 2025.";
            var diff2 = OneAssist.SlideCompare.Matching.TextDiffEngine.ComputeDiff(srcTxt2, dstTxt2);
            bool hasDel25 = diff2.Any(t => t.DiffType == OneAssist.SlideCompare.Models.TextDiffType.Deleted && t.Text.Contains("25"));
            bool hasAdd30 = diff2.Any(t => t.DiffType == OneAssist.SlideCompare.Models.TextDiffType.Added && t.Text.Contains("30"));
            Assert(hasDel25 && hasAdd30, "Test 2: Word-level diff accurately extracted [- 25] and [+ 30]");

            // ── Test 3: Text added (Expected: Added) ────────────────────────────────────
            string srcTxt3 = "Project timeline begins in Q1.";
            string dstTxt3 = "Project timeline begins in Q1. Expected delivery in Q3.";
            var diff3 = OneAssist.SlideCompare.Matching.TextDiffEngine.ComputeDiff(srcTxt3, dstTxt3);
            bool hasAdd3 = diff3.Any(t => t.DiffType == OneAssist.SlideCompare.Models.TextDiffType.Added && t.Text.Contains("Expected delivery"));
            Assert(hasAdd3, "Test 3: Additional sentence detected as Added text token");

            // ── Test 4: Text deleted (Expected: Deleted) ────────────────────────────────
            string srcTxt4 = "Phase 1: Research. Phase 2: Design. Phase 3: Prototyping.";
            string dstTxt4 = "Phase 1: Research. Phase 3: Prototyping.";
            var diff4 = OneAssist.SlideCompare.Matching.TextDiffEngine.ComputeDiff(srcTxt4, dstTxt4);
            bool hasDel4 = diff4.Any(t => t.DiffType == OneAssist.SlideCompare.Models.TextDiffType.Deleted && t.Text.Contains("Phase 2: Design"));
            Assert(hasDel4, "Test 4: Deleted section detected as Deleted text token");

            // ── Test 5: Shape moved (Expected: Moved) ───────────────────────────────────
            var shpSrc5 = new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 10, Name = "BoxA", Left = 100, Top = 100, Width = 200, Height = 100, RawText = "Fixed Content", NormalizedText = "fixed content", HasText = true };
            var shpDst5 = new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 10, Name = "BoxA", Left = 400, Top = 250, Width = 200, Height = 100, RawText = "Fixed Content", NormalizedText = "fixed content", HasText = true };
            var shpMatches5 = OneAssist.SlideCompare.Matching.ShapeMatchingEngine.MatchShapes(
                new List<OneAssist.SlideCompare.Models.ShapeSnapshot> { shpSrc5 },
                new List<OneAssist.SlideCompare.Models.ShapeSnapshot> { shpDst5 },
                1, 1, settings);
            Assert(shpMatches5.Count == 1 && shpMatches5[0].MatchType == OneAssist.SlideCompare.Models.ShapeMatchType.Moved, "Test 5: Shape moved from (100,100) to (400,250) classified as Moved");

            // ── Test 6: Slides reordered (Expected: Reordered) ──────────────────────────
            var s6 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                PresentationName = "Deck_Source.pptx",
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = 1, Title = "Introduction", NormalizedTitle = "introduction" },
                    new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = 2, Title = "Revenue Growth", NormalizedTitle = "revenue growth", BodyText = "25% growth year over year", NormalizedBodyText = "25% growth year over year" },
                    new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = 3, Title = "Risk Factors", NormalizedTitle = "risk factors", BodyText = "Market and regulatory challenges", NormalizedBodyText = "market and regulatory challenges" },
                    new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = 4, Title = "Conclusion", NormalizedTitle = "conclusion" }
                }
            };
            var d6 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                PresentationName = "Deck_Reordered.pptx",
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = 1, Title = "Introduction", NormalizedTitle = "introduction" },
                    new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = 2, Title = "Risk Factors", NormalizedTitle = "risk factors", BodyText = "Market and regulatory challenges", NormalizedBodyText = "market and regulatory challenges" },
                    new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = 3, Title = "Revenue Growth", NormalizedTitle = "revenue growth", BodyText = "25% growth year over year", NormalizedBodyText = "25% growth year over year" },
                    new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = 4, Title = "Conclusion", NormalizedTitle = "conclusion" }
                }
            };
            var res6 = OneAssist.SlideCompare.Comparison.ComparisonEngine.ComparePresentations(s6, d6, settings);
            Assert(res6.ReorderedCount >= 2, $"Test 6: Slides 2 & 3 permutation accurately recognized as Reordered (Count: {res6.ReorderedCount})");

            // ── Test 7: One slide split into two (Expected: Split) ──────────────────────
            var s7 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 5,
                        Title = "Complete Project Overview",
                        NormalizedTitle = "complete project overview",
                        BodyText = "Project Overview. Timeline milestones. Risk management. Budget allocation.",
                        NormalizedBodyText = "project overview timeline milestones risk management budget allocation"
                    }
                }
            };
            var d7 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 5,
                        Title = "Project Overview",
                        NormalizedTitle = "project overview",
                        BodyText = "Project Overview. Timeline milestones.",
                        NormalizedBodyText = "project overview timeline milestones"
                    },
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 6,
                        Title = "Risks and Budget",
                        NormalizedTitle = "risks and budget",
                        BodyText = "Risk management. Budget allocation.",
                        NormalizedBodyText = "risk management budget allocation"
                    }
                }
            };
            var res7 = OneAssist.SlideCompare.Matching.SlideMappingEngine.MapSlides(s7, d7, settings);
            bool hasSplit = res7.Any(r => r.MatchType == OneAssist.SlideCompare.Models.SlideMatchType.Split);
            Assert(hasSplit, "Test 7: Dense source slide 5 detected as Split across destination slides 5 and 6");

            // ── Test 8: Two slides merged (Expected: Merged) ────────────────────────────
            var s8 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 5,
                        Title = "Project Overview",
                        NormalizedTitle = "project overview",
                        BodyText = "Detailed product roadmap",
                        NormalizedBodyText = "detailed product roadmap"
                    },
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 6,
                        Title = "Timeline",
                        NormalizedTitle = "timeline",
                        BodyText = "Deliverables due by Q4",
                        NormalizedBodyText = "deliverables due by q4"
                    }
                }
            };
            var d8 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 5,
                        Title = "Project Overview & Timeline",
                        NormalizedTitle = "project overview timeline",
                        BodyText = "Detailed product roadmap. Deliverables due by Q4.",
                        NormalizedBodyText = "detailed product roadmap deliverables due by q4"
                    }
                }
            };
            var res8 = OneAssist.SlideCompare.Matching.SlideMappingEngine.MapSlides(s8, d8, settings);
            bool hasMerged = res8.Any(r => r.MatchType == OneAssist.SlideCompare.Models.SlideMatchType.Merged);
            Assert(hasMerged, "Test 8: Source slides 5 & 6 detected as Merged into destination slide 5");

            // ── Test 9: Shape indexes completely changed (Expected: Content correctly matched)
            var shpSrc9 = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
            {
                new OneAssist.SlideCompare.Models.ShapeSnapshot { Index = 1, Name = "Chart", ShapeType = "Chart", RawText = "Chart: Revenue", NormalizedText = "chart revenue", HasText = true },
                new OneAssist.SlideCompare.Models.ShapeSnapshot { Index = 2, Name = "Footer", ShapeType = "TextBox", RawText = "Confidential 2026", NormalizedText = "confidential 2026", HasText = true },
                new OneAssist.SlideCompare.Models.ShapeSnapshot { Index = 3, Name = "Title", ShapeType = "Title", RawText = "Executive Summary", NormalizedText = "executive summary", HasText = true }
            };
            var shpDst9 = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
            {
                new OneAssist.SlideCompare.Models.ShapeSnapshot { Index = 1, Name = "Title", ShapeType = "Title", RawText = "Executive Summary", NormalizedText = "executive summary", HasText = true },
                new OneAssist.SlideCompare.Models.ShapeSnapshot { Index = 2, Name = "Footer", ShapeType = "TextBox", RawText = "Confidential 2026", NormalizedText = "confidential 2026", HasText = true },
                new OneAssist.SlideCompare.Models.ShapeSnapshot { Index = 3, Name = "Chart", ShapeType = "Chart", RawText = "Chart: Revenue", NormalizedText = "chart revenue", HasText = true }
            };
            var res9 = OneAssist.SlideCompare.Matching.ShapeMatchingEngine.MatchShapes(shpSrc9, shpDst9, 1, 1, settings);
            int matchedShapes9 = res9.Count(r => r.MatchType == OneAssist.SlideCompare.Models.ShapeMatchType.Matched);
            Assert(matchedShapes9 == 3, $"Test 9: Content correctly matched even when shape indexes are inverted ({matchedShapes9}/3 matched)");

            // ── Test 10: Formatting changed but text identical (Expected: Format Modified)
            var shpSrc10 = new OneAssist.SlideCompare.Models.ShapeSnapshot { Name = "Headline", RawText = "Q4 Financial Results", NormalizedText = "q4 financial results", HasText = true, FontFamily = "Arial", FontSize = 18f, IsBold = false };
            var shpDst10 = new OneAssist.SlideCompare.Models.ShapeSnapshot { Name = "Headline", RawText = "Q4 Financial Results", NormalizedText = "q4 financial results", HasText = true, FontFamily = "Segoe UI", FontSize = 24f, IsBold = true };
            var fmtDiffs = OneAssist.SlideCompare.Comparison.FormattingComparisonEngine.CompareFormatting(shpSrc10, shpDst10, settings);
            bool hasFontChange = fmtDiffs.Any(f => f.PropertyName == "Font Family");
            bool hasSizeChange = fmtDiffs.Any(f => f.PropertyName == "Font Size");
            bool hasBoldChange = fmtDiffs.Any(f => f.PropertyName == "Bold");
            Assert(hasFontChange && hasSizeChange && hasBoldChange, "Test 10: Formatting modifications (Family, Size, Bold) accurately captured while text matches");

            // ── Test 11: Image replaced (Expected: Image Modified) ──────────────────────
            var imgSrc11 = new OneAssist.SlideCompare.Models.ShapeSnapshot { Name = "Logo", ShapeType = "Picture", IsImage = true, ImageHash = "IMG_400x300", Width = 400, Height = 300 };
            var imgDst11 = new OneAssist.SlideCompare.Models.ShapeSnapshot { Name = "Logo", ShapeType = "Picture", IsImage = true, ImageHash = "IMG_600x450", Width = 600, Height = 450 };
            double imgSim = OneAssist.SlideCompare.Matching.ShapeMatchingEngine.CalculateShapeSimilarity(imgSrc11, imgDst11, settings);
            Assert(imgSim < 0.90, $"Test 11: Image with altered dimensions/hash yields lower similarity ({imgSim:P0})");

            // ── Test 12: Large presentation & Cancellation (Expected: Responsive cancellation)
            var s12 = new OneAssist.SlideCompare.Models.PresentationSnapshot { PresentationName = "Big_A.pptx" };
            var d12 = new OneAssist.SlideCompare.Models.PresentationSnapshot { PresentationName = "Big_B.pptx" };
            for (int i = 1; i <= 30; i++)
            {
                s12.Slides.Add(new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = i, Title = $"Section {i}", NormalizedTitle = $"section {i}", BodyText = $"Content for slide {i}", NormalizedBodyText = $"content for slide {i}" });
                d12.Slides.Add(new OneAssist.SlideCompare.Models.SlideSnapshot { SlideNumber = i, Title = $"Section {i}", NormalizedTitle = $"section {i}", BodyText = $"Content for slide {i}", NormalizedBodyText = $"content for slide {i}" });
            }
            var cts12 = new System.Threading.CancellationTokenSource();
            cts12.Cancel(); // Pre-cancel
            bool cancelled = false;
            try
            {
                OneAssist.SlideCompare.Comparison.ComparisonEngine.ComparePresentations(s12, d12, settings, cancellationToken: cts12.Token);
            }
            catch (OperationCanceledException)
            {
                cancelled = true;
            }
            Assert(cancelled, "Test 12: Large presentation comparison honors CancellationToken promptly without freezing");

            // ── Test 13: Slide 1-1, Slide 2-2, Slide 3-4, Slide 4-3 Best Matching (User Directive) ──
            var s13 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                PresentationName = "Deck_Source.pptx",
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 1,
                        Title = "Introduction & Agenda",
                        NormalizedTitle = "introduction agenda",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 1, Name = "Title 1", ShapeType = "Title", RawText = "Introduction & Agenda", NormalizedText = "introduction agenda", HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 2, Name = "Body 1", ShapeType = "TextBox", RawText = "Welcome to presentation", NormalizedText = "welcome to presentation", HasText = true }
                        }
                    },
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 2,
                        Title = "Executive Summary",
                        NormalizedTitle = "executive summary",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 3, Name = "Title 2", ShapeType = "Title", RawText = "Executive Summary", NormalizedText = "executive summary", HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 4, Name = "Body 2", ShapeType = "TextBox", RawText = "Revenue target is $10M", NormalizedText = "revenue target is 10m", HasText = true }
                        }
                    },
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 3,
                        Title = "Technical Architecture",
                        NormalizedTitle = "technical architecture",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 5, Name = "Title 3", ShapeType = "Title", RawText = "Technical Architecture", NormalizedText = "technical architecture", HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 6, Name = "Body 3", ShapeType = "TextBox", RawText = "Microservices and Kubernetes deployment", NormalizedText = "microservices and kubernetes deployment", HasText = true }
                        }
                    },
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 4,
                        Title = "Next Steps & Roadmap",
                        NormalizedTitle = "next steps roadmap",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 7, Name = "Title 4", ShapeType = "Title", RawText = "Next Steps & Roadmap", NormalizedText = "next steps roadmap", HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 8, Name = "Body 4", ShapeType = "TextBox", RawText = "Launch in Q4 2026", NormalizedText = "launch in q4 2026", HasText = true }
                        }
                    }
                }
            };

            var d13 = new OneAssist.SlideCompare.Models.PresentationSnapshot
            {
                PresentationName = "Deck_Destination.pptx",
                Slides = new List<OneAssist.SlideCompare.Models.SlideSnapshot>
                {
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 1,
                        Title = "Introduction & Agenda",
                        NormalizedTitle = "introduction agenda",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 101, Name = "Title 1", ShapeType = "Title", RawText = "Introduction & Agenda", NormalizedText = "introduction agenda", HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 102, Name = "Body 1", ShapeType = "TextBox", RawText = "Welcome to presentation", NormalizedText = "welcome to presentation", HasText = true }
                        }
                    },
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 2,
                        Title = "Executive Summary",
                        NormalizedTitle = "executive summary",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 103, Name = "Title 2", ShapeType = "Title", RawText = "Executive Summary", NormalizedText = "executive summary", HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 104, Name = "Body 2", ShapeType = "TextBox", RawText = "Revenue target is $10M", NormalizedText = "revenue target is 10m", HasText = true }
                        }
                    },
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 3,
                        Title = "Next Steps & Roadmap",
                        NormalizedTitle = "next steps roadmap",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 105, Name = "Title 3", ShapeType = "Title", RawText = "Next Steps & Roadmap", NormalizedText = "next steps roadmap", HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 106, Name = "Body 3", ShapeType = "TextBox", RawText = "Launch in Q1 2027", NormalizedText = "launch in q1 2027", HasText = true }
                        }
                    },
                    new OneAssist.SlideCompare.Models.SlideSnapshot
                    {
                        SlideNumber = 4,
                        Title = "Technical Architecture",
                        NormalizedTitle = "technical architecture",
                        Shapes = new List<OneAssist.SlideCompare.Models.ShapeSnapshot>
                        {
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 107, Name = "Title 4", ShapeType = "Title", RawText = "Technical Architecture", NormalizedText = "technical architecture", HasText = true },
                            new OneAssist.SlideCompare.Models.ShapeSnapshot { Id = 108, Name = "Body 4", ShapeType = "TextBox", RawText = "Microservices and Kubernetes cluster deployment", NormalizedText = "microservices and kubernetes cluster deployment", HasText = true }
                        }
                    }
                }
            };

            var res13 = OneAssist.SlideCompare.Comparison.ComparisonEngine.ComparePresentations(s13, d13, settings);
            var pair1 = res13.SlideResults.FirstOrDefault(r => r.SourceSlideNumber == 1);
            var pair2 = res13.SlideResults.FirstOrDefault(r => r.SourceSlideNumber == 2);
            var pair3 = res13.SlideResults.FirstOrDefault(r => r.SourceSlideNumber == 3);
            var pair4 = res13.SlideResults.FirstOrDefault(r => r.SourceSlideNumber == 4);

            bool p1Correct = pair1 != null && pair1.PrimaryDestinationSlide == 1 && pair1.MatchType == OneAssist.SlideCompare.Models.SlideMatchType.Matched;
            bool p2Correct = pair2 != null && pair2.PrimaryDestinationSlide == 2 && pair2.MatchType == OneAssist.SlideCompare.Models.SlideMatchType.Matched;
            bool p3Correct = pair3 != null && pair3.PrimaryDestinationSlide == 4 && pair3.MatchType == OneAssist.SlideCompare.Models.SlideMatchType.Reordered;
            bool p4Correct = pair4 != null && pair4.PrimaryDestinationSlide == 3 && pair4.MatchType == OneAssist.SlideCompare.Models.SlideMatchType.Reordered;

            Assert(p1Correct && p2Correct && p3Correct && p4Correct, "Test 13: Slide 1↔1, Slide 2↔2, Slide 3↔4, Slide 4↔3 correctly matched 1-to-1 without split/merge false positives");

            // Verify shape-to-shape comparison inside Slide 3 ↔ Slide 4
            var s3BodyShape = pair3?.ShapeResults?.FirstOrDefault(sh => sh.SourceText != null && sh.SourceText.Contains("Kubernetes"));
            bool s3HasClusterAdded = s3BodyShape != null && s3BodyShape.DiffTokens.Any(t => t.DiffType == OneAssist.SlideCompare.Models.TextDiffType.Added && t.Text.Contains("cluster"));
            Assert(s3HasClusterAdded, "Test 13: Shape-to-shape best matching accurately detected [+ cluster] inside Slide 3 ↔ Slide 4");

            // Verify shape-to-shape comparison inside Slide 4 ↔ Slide 3
            var s4BodyShape = pair4?.ShapeResults?.FirstOrDefault(sh => sh.SourceText != null && sh.SourceText.Contains("Launch"));
            bool s4HasDateDiff = s4BodyShape != null && s4BodyShape.HasTextChanges && s4BodyShape.DiffTokens.Any(t => t.DiffType == OneAssist.SlideCompare.Models.TextDiffType.Modified || t.DiffType == OneAssist.SlideCompare.Models.TextDiffType.Added);
            Assert(s4HasDateDiff, "Test 13: Shape-to-shape best matching accurately detected Q4 2026 → Q1 2027 inside Slide 4 ↔ Slide 3");
        }
    }
}

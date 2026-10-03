using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OneAssist.Assistant.AILayout.Models;
using OneAssist.Assistant.Services;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.AILayout.Analysis
{
    public static class SlideSemanticAnalyzer
    {
        public static SlideSemanticAnalysis Analyze(PowerPoint.Slide slide)
        {
            var analysis = new SlideSemanticAnalysis();
            if (slide == null) return analysis;

            try
            {
                analysis.SlideIndex = slide.SlideIndex;
                var pres = slide.Parent as PowerPoint.Presentation;
                if (pres != null)
                {
                    analysis.SlideWidth = pres.PageSetup.SlideWidth;
                    analysis.SlideHeight = pres.PageSetup.SlideHeight;
                }

                // 1. Extract and classify all shapes
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    var semShape = ClassifyShape(s, analysis.SlideWidth, analysis.SlideHeight);
                    if (semShape != null)
                    {
                        analysis.SemanticShapes.Add(semShape);
                    }
                }

                // Identify Slide Title
                var title = analysis.SemanticShapes.FirstOrDefault(s => s.SemanticType == ObjectSemanticType.Title);
                if (title != null)
                {
                    analysis.SlideTitle = title.RawText.Trim();
                }
                else
                {
                    // Fallback to highest shape on slide with text
                    var topText = analysis.SemanticShapes
                        .Where(s => !string.IsNullOrWhiteSpace(s.RawText) && s.Top < analysis.SlideHeight * 0.4f)
                        .OrderBy(s => s.Top)
                        .ThenByDescending(s => s.FontSize)
                        .FirstOrDefault();

                    if (topText != null)
                    {
                        topText.SemanticType = ObjectSemanticType.Title;
                        analysis.SlideTitle = topText.RawText.Trim();
                    }
                }

                // 2. Classify Slide Content Type (Req 2 & 3)
                ClassifySlideContentType(analysis);

                // 3. Build Detected Structure Summary (Req 1)
                BuildStructureSummary(analysis);

                // 4. Detect Visual Hierarchy and Density Issues (Req 23 & 24)
                DetectLayoutIssues(analysis);
            }
            catch { }

            return analysis;
        }

        public static SemanticShape ClassifyShape(PowerPoint.Shape shape, float slideWidth, float slideHeight)
        {
            if (shape == null) return null;

            var sem = new SemanticShape
            {
                ShapeId = shape.Id,
                ShapeName = shape.Name,
                Left = shape.Left,
                Top = shape.Top,
                Width = shape.Width,
                Height = shape.Height,
                ZOrderPosition = shape.ZOrderPosition
            };

            // Check Chart
            if (shape.HasChart == Office.MsoTriState.msoTrue)
            {
                sem.HasChart = true;
                sem.SemanticType = (shape.Width * shape.Height > (slideWidth * slideHeight * 0.25f))
                    ? ObjectSemanticType.PrimaryChart
                    : ObjectSemanticType.SupportingChart;
                return sem;
            }

            // Check Table
            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                sem.HasTable = true;
                sem.SemanticType = ObjectSemanticType.Table;
                return sem;
            }

            // Check Picture
            if (shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture)
            {
                sem.HasPicture = true;
                float areaRatio = (shape.Width * shape.Height) / (slideWidth * slideHeight);

                if (areaRatio > 0.30f)
                {
                    sem.SemanticType = ObjectSemanticType.HeroImage;
                }
                else if (shape.Width <= 80f && shape.Height <= 80f && (shape.Left < 80f || shape.Left > slideWidth - 120f))
                {
                    sem.SemanticType = ObjectSemanticType.Logo;
                }
                else if (shape.Width <= 50f && shape.Height <= 50f)
                {
                    sem.SemanticType = ObjectSemanticType.Icon;
                }
                else
                {
                    sem.SemanticType = ObjectSemanticType.SupportingImage;
                }
                return sem;
            }

            // Check Text
            if (SafeInteropHelper.HasText(shape))
            {
                string text = shape.TextFrame.TextRange.Text ?? string.Empty;
                sem.RawText = text;
                var words = text.Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                sem.WordCount = words.Length;

                try
                {
                    sem.FontName = shape.TextFrame.TextRange.Font.Name;
                    sem.FontSize = shape.TextFrame.TextRange.Font.Size;
                    sem.IsBold = shape.TextFrame.TextRange.Font.Bold == Office.MsoTriState.msoTrue;
                }
                catch { }

                // Count bullets
                var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                sem.BulletCount = lines.Count(l => l.TrimStart().StartsWith("•") || l.TrimStart().StartsWith("-") || l.TrimStart().StartsWith("*"));
                if (sem.BulletCount == 0 && lines.Length > 1)
                {
                    sem.BulletCount = lines.Length;
                }

                // Check Placeholder
                if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                {
                    var phType = shape.PlaceholderFormat.Type;
                    if (phType == PowerPoint.PpPlaceholderType.ppPlaceholderTitle || phType == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle)
                    {
                        sem.SemanticType = ObjectSemanticType.Title;
                        return sem;
                    }
                    if (phType == PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle)
                    {
                        sem.SemanticType = ObjectSemanticType.Subtitle;
                        return sem;
                    }
                    if (phType == PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                    {
                        sem.SemanticType = ObjectSemanticType.Body;
                        return sem;
                    }
                }

                // KPI Detection: Short text with numbers, currency, percentage, multipliers
                if (sem.WordCount <= 6 && Regex.IsMatch(text, @"^[\$€£¥\+]?[\d,\.]+\s*(%|[kKmMbBtT]|pts|x|X)?$") || (sem.FontSize >= 26f && sem.WordCount <= 4))
                {
                    sem.SemanticType = ObjectSemanticType.KPI;
                    return sem;
                }

                // Footnote Detection: small font near slide bottom
                if (sem.FontSize <= 10.5f && sem.Top > slideHeight * 0.82f)
                {
                    sem.SemanticType = ObjectSemanticType.Footnote;
                    return sem;
                }

                // Title by font size & position
                if (sem.FontSize >= 22f && sem.Top < slideHeight * 0.35f && sem.WordCount <= 18)
                {
                    sem.SemanticType = ObjectSemanticType.Title;
                    return sem;
                }

                // Subtitle by position right below title
                if (sem.FontSize >= 14f && sem.FontSize < 22f && sem.Top < slideHeight * 0.45f && sem.WordCount <= 25)
                {
                    sem.SemanticType = ObjectSemanticType.Subtitle;
                    return sem;
                }

                // Default text is Body or Description
                sem.SemanticType = sem.BulletCount > 0 ? ObjectSemanticType.Body : ObjectSemanticType.Description;
                return sem;
            }

            // Shape container / card / decoration
            if (shape.Width >= 120f && shape.Height >= 80f)
            {
                sem.SemanticType = ObjectSemanticType.Card;
            }
            else if (shape.Height <= 4f || shape.Width <= 4f)
            {
                sem.SemanticType = ObjectSemanticType.Divider;
            }
            else
            {
                sem.SemanticType = ObjectSemanticType.Decoration;
            }

            return sem;
        }

        private static void ClassifySlideContentType(SlideSemanticAnalysis analysis)
        {
            string titleLower = analysis.SlideTitle.ToLowerInvariant();

            // 1. KPI / Financial
            bool hasKPI = analysis.SemanticShapes.Any(s => s.SemanticType == ObjectSemanticType.KPI);
            bool hasChart = analysis.TotalChartsCount > 0;
            bool hasTable = analysis.TotalTablesCount > 0;

            if (hasKPI && hasChart)
            {
                analysis.DetectedContentType = SlideContentType.KPI;
                analysis.ContentTypeReason = "Slide contains one or more key performance metrics accompanied by supporting chart evidence.";
                return;
            }

            if (hasKPI)
            {
                analysis.DetectedContentType = SlideContentType.KPI;
                analysis.ContentTypeReason = "Slide prominently showcases core numerical metrics / headline statistics.";
                return;
            }

            if (hasChart)
            {
                analysis.DetectedContentType = SlideContentType.Chart;
                analysis.ContentTypeReason = "Data chart is the focal point of the visual presentation.";
                return;
            }

            if (hasTable)
            {
                analysis.DetectedContentType = SlideContentType.Table;
                analysis.ContentTypeReason = "Tabular data structure detected.";
                return;
            }

            // 2. Keyword & Structure Analysis
            if (titleLower.Contains("timeline") || titleLower.Contains("milestone") || titleLower.Contains("roadmap"))
            {
                analysis.DetectedContentType = SlideContentType.Timeline;
                analysis.ContentTypeReason = "Time-based roadmap/milestone sequence detected in title and layout.";
                return;
            }

            if (titleLower.Contains("process") || titleLower.Contains("workflow") || titleLower.Contains("steps") || titleLower.Contains("journey"))
            {
                analysis.DetectedContentType = SlideContentType.Process;
                analysis.ContentTypeReason = "Sequential process flow or workflow steps detected.";
                return;
            }

            if (titleLower.Contains("versus") || titleLower.Contains(" vs ") || titleLower.Contains("comparison") || titleLower.Contains("before") || titleLower.Contains("pros and cons"))
            {
                analysis.DetectedContentType = SlideContentType.Comparison;
                analysis.ContentTypeReason = "Side-by-side comparative analysis structure detected.";
                return;
            }

            if (titleLower.Contains("strategy") || titleLower.Contains("pillars") || titleLower.Contains("priorities") || titleLower.Contains("framework"))
            {
                analysis.DetectedContentType = SlideContentType.Strategy;
                analysis.ContentTypeReason = "Strategic pillars and framework overview detected.";
                return;
            }

            if (titleLower.Contains("executive summary") || titleLower.Contains("key takeaways") || titleLower.Contains("overview"))
            {
                analysis.DetectedContentType = SlideContentType.ExecutiveSummary;
                analysis.ContentTypeReason = "Executive summary / core takeaways detected.";
                return;
            }

            if (titleLower.Contains("agenda") || titleLower.Contains("table of contents"))
            {
                analysis.DetectedContentType = SlideContentType.Agenda;
                analysis.ContentTypeReason = "Meeting agenda or presentation structure.";
                return;
            }

            // 3. Density-driven classification
            if (analysis.TotalWordCount >= 110 || analysis.TotalBulletCount >= 7)
            {
                analysis.DetectedContentType = SlideContentType.TextHeavy;
                analysis.ContentTypeReason = $"High text volume ({analysis.TotalWordCount} words, {analysis.TotalBulletCount} bullets) requiring structured card partitioning.";
                return;
            }

            if (analysis.TotalImagesCount >= 2 || (analysis.TotalImagesCount == 1 && analysis.TotalWordCount < 30))
            {
                analysis.DetectedContentType = SlideContentType.ImageHeavy;
                analysis.ContentTypeReason = "Visual media/image storytelling is the primary emphasis.";
                return;
            }

            // Fallback
            analysis.DetectedContentType = SlideContentType.GeneralContent;
            analysis.ContentTypeReason = "Standard informative slide layout.";
        }

        private static void BuildStructureSummary(SlideSemanticAnalysis analysis)
        {
            if (!string.IsNullOrWhiteSpace(analysis.SlideTitle))
            {
                analysis.DetectedStructureSummary.Add("Title Header");
            }
            if (analysis.SemanticShapes.Any(s => s.SemanticType == ObjectSemanticType.Subtitle))
            {
                analysis.DetectedStructureSummary.Add("Context Subtitle");
            }
            if (analysis.KPIShapes.Count > 0)
            {
                analysis.DetectedStructureSummary.Add($"{analysis.KPIShapes.Count} KPI Metric(s)");
            }
            if (analysis.TotalChartsCount > 0)
            {
                analysis.DetectedStructureSummary.Add("Data Chart Visual");
            }
            if (analysis.TotalTablesCount > 0)
            {
                analysis.DetectedStructureSummary.Add("Structured Table");
            }
            if (analysis.BodyShapes.Count > 0)
            {
                analysis.DetectedStructureSummary.Add($"{analysis.TotalBulletCount} Bullet Point(s) / Text Block(s)");
            }
            if (analysis.TotalImagesCount > 0)
            {
                analysis.DetectedStructureSummary.Add($"{analysis.TotalImagesCount} Image Asset(s)");
            }
        }

        private static void DetectLayoutIssues(SlideSemanticAnalysis analysis)
        {
            if (analysis.TotalWordCount > 130)
            {
                analysis.DetectedIssues.Add($"High text density ({analysis.TotalWordCount} words). A structured card layout is recommended.");
            }
            if (analysis.TotalBulletCount >= 6)
            {
                analysis.DetectedIssues.Add($"Excessive bullet points ({analysis.TotalBulletCount} bullets). Multi-column cards improve visual scanability.");
            }

            var title = analysis.TitleShape;
            var bodies = analysis.BodyShapes;
            if (title != null && bodies.Any(b => b.FontSize >= title.FontSize - 3f))
            {
                analysis.DetectedIssues.Add("Weak visual hierarchy: body text font size is nearly equal to title font size.");
            }

            if (analysis.TotalShapesCount >= 8 && analysis.DetectedContentType != SlideContentType.Process)
            {
                analysis.DetectedIssues.Add("High object fragmentation. Aligning onto a structured 12-column grid will reduce cognitive load.");
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using OneAssist.SlideCompare.Matching;
using OneAssist.SlideCompare.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneAssist.SlideCompare.Extraction
{
    /// <summary>
    /// Extracts a complete SlideSnapshot from a PowerPoint COM Slide.
    /// </summary>
    public static class SlideExtractor
    {
        public static SlideSnapshot ExtractSlide(PowerPoint.Slide slide, int slideNumber)
        {
            if (slide == null) return null;

            var snapshot = new SlideSnapshot
            {
                SlideNumber = slideNumber
            };

            try { snapshot.SlideId = slide.SlideID; } catch { snapshot.SlideId = slideNumber; }

            // Hidden slide check (Req 30)
            try
            {
                snapshot.IsHidden = slide.SlideShowTransition.Hidden == Office.MsoTriState.msoTrue;
            }
            catch { }

            // Speaker notes check (Req 29)
            try
            {
                if (slide.NotesPage != null && slide.NotesPage.Shapes != null)
                {
                    int noteShapes = slide.NotesPage.Shapes.Count;
                    for (int n = 1; n <= noteShapes; n++)
                    {
                        var shp = slide.NotesPage.Shapes[n];
                        if (shp.Type == Office.MsoShapeType.msoPlaceholder && shp.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        {
                            snapshot.SpeakerNotes = TextExtractor.ExtractText(shp);
                            break;
                        }
                    }
                }
            }
            catch { }

            // Extract shapes
            var bodyTexts = new List<string>();
            string titleCandidate = string.Empty;

            try
            {
                int shapeCount = slide.Shapes.Count;
                for (int i = 1; i <= shapeCount; i++)
                {
                    try
                    {
                        var shp = slide.Shapes[i];
                        var shpSnapshot = ShapeExtractor.ExtractShape(shp, i);
                        if (shpSnapshot != null)
                        {
                            snapshot.Shapes.Add(shpSnapshot);

                            // Detect title
                            if (shpSnapshot.IsPlaceholder && shpSnapshot.PlaceholderType.IndexOf("Title", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                if (string.IsNullOrEmpty(titleCandidate) && !string.IsNullOrWhiteSpace(shpSnapshot.RawText))
                                {
                                    titleCandidate = shpSnapshot.RawText.Trim();
                                }
                            }
                            else if (string.IsNullOrEmpty(titleCandidate) && shpSnapshot.HasText && shpSnapshot.FontSize >= 24)
                            {
                                titleCandidate = shpSnapshot.RawText.Trim();
                            }
                            else if (shpSnapshot.HasText)
                            {
                                bodyTexts.Add(shpSnapshot.RawText.Trim());
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            snapshot.Title = titleCandidate;
            snapshot.NormalizedTitle = SimilarityCalculator.NormalizeText(titleCandidate);

            snapshot.BodyText = string.Join("\n", bodyTexts);
            snapshot.NormalizedBodyText = SimilarityCalculator.NormalizeText(snapshot.BodyText);

            // Compute composite fingerprint
            snapshot.Fingerprint = ContentMatchingEngine.GenerateSlideFingerprint(snapshot);

            return snapshot;
        }
    }
}

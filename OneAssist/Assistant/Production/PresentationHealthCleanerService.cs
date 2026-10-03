using System;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Production
{
    public class CleanerSummary
    {
        public int CommentsRemoved { get; set; }
        public int HiddenSlidesRemoved { get; set; }
        public int NotesCleared { get; set; }
        public int GhostShapesRemoved { get; set; }
    }

    public static class PresentationHealthCleanerService
    {
        public static CleanerSummary CleanPresentation(PowerPoint.Presentation pres, bool removeComments, bool removeHiddenSlides, bool clearSpeakerNotes, bool removeGhostShapes)
        {
            var summary = new CleanerSummary();
            if (pres == null) return summary;

            try
            {
                // 1. Comments
                if (removeComments)
                {
                    foreach (PowerPoint.Slide s in pres.Slides)
                    {
                        try
                        {
                            while (s.Comments.Count > 0)
                            {
                                s.Comments[1].Delete();
                                summary.CommentsRemoved++;
                            }
                        }
                        catch { }
                    }
                }

                // 2. Speaker Notes
                if (clearSpeakerNotes)
                {
                    foreach (PowerPoint.Slide s in pres.Slides)
                    {
                        try
                        {
                            if (s.NotesPage != null && s.NotesPage.Shapes.Count > 0)
                            {
                                foreach (PowerPoint.Shape ns in s.NotesPage.Shapes)
                                {
                                    if (ns.PlaceholderFormat != null && ns.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                                    {
                                        if (ns.HasTextFrame == Office.MsoTriState.msoTrue && !string.IsNullOrWhiteSpace(ns.TextFrame.TextRange.Text))
                                        {
                                            ns.TextFrame.TextRange.Text = string.Empty;
                                            summary.NotesCleared++;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }

                // 3. Ghost shapes
                if (removeGhostShapes)
                {
                    foreach (PowerPoint.Slide s in pres.Slides)
                    {
                        var candidates = Services.SlideCleanupService.FindCleanupCandidates(s, s.SlideIndex);
                        foreach (var c in candidates)
                        {
                            summary.GhostShapesRemoved += Services.SlideCleanupService.DeleteCandidate(s, c.ShapeId);
                        }
                    }
                }

                // 4. Hidden Slides
                if (removeHiddenSlides)
                {
                    for (int i = pres.Slides.Count; i >= 1; i--)
                    {
                        try
                        {
                            if (pres.Slides[i].SlideShowTransition.Hidden == Office.MsoTriState.msoTrue)
                            {
                                pres.Slides[i].Delete();
                                summary.HiddenSlidesRemoved++;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            return summary;
        }
    }
}

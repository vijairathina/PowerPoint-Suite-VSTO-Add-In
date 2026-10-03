using System;
using System.Collections.Generic;
using OneAssist.Assistant.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Production
{
    public class SearchMatchItem
    {
        public int SlideIndex { get; set; }
        public int ShapeId { get; set; }
        public string ShapeName { get; set; }
        public string MatchType { get; set; } // Text, Font, ShapeName, Color
        public string MatchSnippet { get; set; }
    }

    public static class PresentationSearchService
    {
        public static List<SearchMatchItem> SearchPresentation(PowerPoint.Presentation pres, string query)
        {
            var results = new List<SearchMatchItem>();
            if (pres == null || string.IsNullOrWhiteSpace(query)) return results;

            string q = query.Trim();

            try
            {
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        // 1. Text Content
                        if (SafeInteropHelper.HasText(shape))
                        {
                            string text = shape.TextFrame.TextRange.Text;
                            if (!string.IsNullOrEmpty(text) && text.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                results.Add(new SearchMatchItem
                                {
                                    SlideIndex = slide.SlideIndex,
                                    ShapeId = shape.Id,
                                    ShapeName = shape.Name,
                                    MatchType = "Text",
                                    MatchSnippet = text.Length > 50 ? text.Substring(0, 47) + "..." : text
                                });
                            }

                            // 2. Font name
                            string fn = SafeInteropHelper.GetFontName(shape);
                            if (!string.IsNullOrEmpty(fn) && fn.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                results.Add(new SearchMatchItem
                                {
                                    SlideIndex = slide.SlideIndex,
                                    ShapeId = shape.Id,
                                    ShapeName = shape.Name,
                                    MatchType = "Font",
                                    MatchSnippet = $"Font: {fn}"
                                });
                            }
                        }

                        // 3. Shape name
                        if (!string.IsNullOrEmpty(shape.Name) && shape.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            results.Add(new SearchMatchItem
                            {
                                SlideIndex = slide.SlideIndex,
                                ShapeId = shape.Id,
                                ShapeName = shape.Name,
                                MatchType = "ShapeName",
                                MatchSnippet = $"Name: {shape.Name}"
                            });
                        }
                    }
                }
            }
            catch { }

            return results;
        }
    }
}

using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Layers
{
    public class MorphMatchCandidate
    {
        public int Slide1Index { get; set; }
        public int Shape1Id { get; set; }
        public string Shape1Name { get; set; }

        public int Slide2Index { get; set; }
        public int Shape2Id { get; set; }
        public string Shape2Name { get; set; }

        public string ProposedMorphName { get; set; }
    }

    public static class MorphAssistantService
    {
        public static List<MorphMatchCandidate> DetectCrossSlideMorphMatches(PowerPoint.Slide slide1, PowerPoint.Slide slide2)
        {
            var matches = new List<MorphMatchCandidate>();
            if (slide1 == null || slide2 == null) return matches;

            try
            {
                foreach (PowerPoint.Shape s1 in slide1.Shapes)
                {
                    foreach (PowerPoint.Shape s2 in slide2.Shapes)
                    {
                        // Check if shapes share same type and similar aspect ratio
                        if (s1.Type == s2.Type)
                        {
                            string cleanName = s1.Name.TrimStart('!');
                            matches.Add(new MorphMatchCandidate
                            {
                                Slide1Index = slide1.SlideIndex,
                                Shape1Id = s1.Id,
                                Shape1Name = s1.Name,
                                Slide2Index = slide2.SlideIndex,
                                Shape2Id = s2.Id,
                                Shape2Name = s2.Name,
                                ProposedMorphName = $"!!{cleanName}"
                            });
                            break;
                        }
                    }
                }
            }
            catch { }

            return matches;
        }

        public static void ApplyMorphMatching(PowerPoint.Slide slide1, int shape1Id, PowerPoint.Slide slide2, int shape2Id, string morphName)
        {
            if (slide1 == null || slide2 == null || string.IsNullOrWhiteSpace(morphName)) return;

            string taggedName = morphName.StartsWith("!!") ? morphName : $"!!{morphName}";

            try
            {
                foreach (PowerPoint.Shape s in slide1.Shapes)
                {
                    if (s.Id == shape1Id)
                    {
                        s.Name = taggedName;
                        break;
                    }
                }
                foreach (PowerPoint.Shape s in slide2.Shapes)
                {
                    if (s.Id == shape2Id)
                    {
                        s.Name = taggedName;
                        break;
                    }
                }
            }
            catch { }
        }
    }
}

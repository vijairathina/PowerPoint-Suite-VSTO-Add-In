using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.AILayout.Analysis;
using OneAssist.Assistant.AILayout.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.AILayout.Engine
{
    public static class ReferenceSlideEngine
    {
        public static LayoutSpecification AdaptReferenceSlideLayout(PowerPoint.Slide referenceSlide, PowerPoint.Slide targetSlide)
        {
            if (referenceSlide == null || targetSlide == null) return null;

            var refAnalysis = SlideSemanticAnalyzer.Analyze(referenceSlide);
            var targetAnalysis = SlideSemanticAnalyzer.Analyze(targetSlide);

            var spec = new LayoutSpecification
            {
                SlideType = refAnalysis.DetectedContentType.ToString(),
                Layout = $"reference_from_slide_{referenceSlide.SlideIndex}",
                ReasoningSummary = $"Adapted layout structure and visual spacing from Slide {referenceSlide.SlideIndex} while preserving Slide {targetSlide.SlideIndex}'s content."
            };

            // 1. Extract layout regions from reference slide
            int cardIdx = 1;
            foreach (var s in refAnalysis.SemanticShapes.OrderBy(s => s.Top).ThenBy(s => s.Left))
            {
                string regionName;
                if (s.SemanticType == ObjectSemanticType.Title)
                    regionName = "Title";
                else if (s.SemanticType == ObjectSemanticType.KPI)
                    regionName = "KPI_Callout";
                else if (s.HasChart)
                    regionName = "Primary_Chart";
                else if (s.HasPicture)
                    regionName = "Hero_Image";
                else
                    regionName = $"Region_Card_{cardIdx++}";

                spec.Regions.Add(new LayoutRegion
                {
                    Name = regionName,
                    X = s.Left,
                    Y = s.Top,
                    Width = s.Width,
                    Height = s.Height,
                    PreferredType = s.SemanticType.ToString(),
                    IsContainer = s.SemanticType == ObjectSemanticType.Card || s.SemanticType == ObjectSemanticType.Container
                });
            }

            // 2. Map target slide's shapes into these extracted regions
            var assignedTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Match Title
            var targetTitle = targetAnalysis.TitleShape;
            var titleRegion = spec.Regions.FirstOrDefault(r => r.Name.Equals("Title", StringComparison.OrdinalIgnoreCase));
            if (targetTitle != null && titleRegion != null)
            {
                spec.ObjectAssignments.Add(new ObjectAssignment
                {
                    ShapeId = targetTitle.ShapeId,
                    TargetRegion = titleRegion.Name,
                    Role = "Title"
                });
                assignedTargets.Add(titleRegion.Name);
            }

            // Match Charts
            var targetCharts = targetAnalysis.ChartShapes;
            var chartRegions = spec.Regions.Where(r => r.PreferredType.Contains("Chart")).ToList();
            for (int i = 0; i < targetCharts.Count && i < chartRegions.Count; i++)
            {
                spec.ObjectAssignments.Add(new ObjectAssignment
                {
                    ShapeId = targetCharts[i].ShapeId,
                    TargetRegion = chartRegions[i].Name,
                    Role = "Chart"
                });
                assignedTargets.Add(chartRegions[i].Name);
            }

            // Match Images
            var targetImages = targetAnalysis.ImageShapes;
            var imgRegions = spec.Regions.Where(r => r.PreferredType.Contains("Image")).ToList();
            for (int i = 0; i < targetImages.Count && i < imgRegions.Count; i++)
            {
                spec.ObjectAssignments.Add(new ObjectAssignment
                {
                    ShapeId = targetImages[i].ShapeId,
                    TargetRegion = imgRegions[i].Name,
                    Role = "Image"
                });
                assignedTargets.Add(imgRegions[i].Name);
            }

            // Match remaining body shapes
            var remainingTargetShapes = targetAnalysis.SemanticShapes
                .Where(s => spec.ObjectAssignments.All(a => a.ShapeId != s.ShapeId))
                .ToList();

            var availableRegions = spec.Regions.Where(r => !assignedTargets.Contains(r.Name)).ToList();

            for (int i = 0; i < remainingTargetShapes.Count; i++)
            {
                var reg = (i < availableRegions.Count) ? availableRegions[i] : availableRegions.LastOrDefault();
                if (reg != null)
                {
                    spec.ObjectAssignments.Add(new ObjectAssignment
                    {
                        ShapeId = remainingTargetShapes[i].ShapeId,
                        TargetRegion = reg.Name,
                        Role = "Body"
                    });
                }
            }

            return spec;
        }
    }
}

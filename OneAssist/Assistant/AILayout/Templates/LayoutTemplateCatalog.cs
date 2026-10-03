using System;
using System.Collections.Generic;
using System.Linq;
using OneAssist.Assistant.AILayout.Models;

namespace OneAssist.Assistant.AILayout.Templates
{
    public static class LayoutTemplateCatalog
    {
        public const float SlideWidth = 960f;  // 13.33 in in points
        public const float SlideHeight = 540f; // 7.5 in in points
        public const float MarginX = 36f;     // 0.5 in
        public const float MarginTop = 36f;    // 0.5 in
        public const float TitleHeight = 44f;
        public const float Gap = 14.4f;        // 0.2 in

        public static LayoutSpecification CreateKPIPlusChartLayout(SlideSemanticAnalysis analysis)
        {
            var spec = new LayoutSpecification
            {
                SlideType = "executive_kpi",
                Layout = "kpi_plus_chart",
                ReasoningSummary = "Showcases the prominent KPI metric as a bold visual anchor on the left with the supporting data chart occupying the primary visual area on the right."
            };

            // Title Region
            spec.Regions.Add(new LayoutRegion
            {
                Name = "Title",
                X = MarginX,
                Y = MarginTop,
                Width = SlideWidth - (2 * MarginX),
                Height = TitleHeight,
                PreferredType = "Title"
            });

            float contentTop = MarginTop + TitleHeight + Gap;
            float contentHeight = SlideHeight - contentTop - MarginTop;

            // KPI Box (Left Column: ~32% width)
            float kpiWidth = 280f;
            spec.Regions.Add(new LayoutRegion
            {
                Name = "KPI_Callout",
                X = MarginX,
                Y = contentTop,
                Width = kpiWidth,
                Height = 160f,
                PreferredType = "KPI",
                IsContainer = true,
                BackgroundColorHex = "#F1F5F9",
                BorderColorHex = "#CBD5E1"
            });

            // Key Takeaways underneath KPI
            spec.Regions.Add(new LayoutRegion
            {
                Name = "KPI_Insights",
                X = MarginX,
                Y = contentTop + 160f + Gap,
                Width = kpiWidth,
                Height = contentHeight - 160f - Gap,
                PreferredType = "Body",
                IsContainer = true,
                BackgroundColorHex = "#FFFFFF",
                BorderColorHex = "#E2E8F0"
            });

            // Chart (Right Column: ~65% width)
            float chartLeft = MarginX + kpiWidth + Gap;
            float chartWidth = SlideWidth - chartLeft - MarginX;
            spec.Regions.Add(new LayoutRegion
            {
                Name = "Primary_Chart",
                X = chartLeft,
                Y = contentTop,
                Width = chartWidth,
                Height = contentHeight,
                PreferredType = "Chart"
            });

            AssignShapesToRegions(spec, analysis);
            return spec;
        }

        public static LayoutSpecification CreateThreeCardsLayout(SlideSemanticAnalysis analysis)
        {
            var spec = new LayoutSpecification
            {
                SlideType = "content_cards",
                Layout = "three_cards",
                ReasoningSummary = "Structures dense information into three balanced, digestible cards, establishing clean visual cadence and eliminating crowded bullet lists."
            };

            // Title
            spec.Regions.Add(new LayoutRegion
            {
                Name = "Title",
                X = MarginX,
                Y = MarginTop,
                Width = SlideWidth - (2 * MarginX),
                Height = TitleHeight,
                PreferredType = "Title"
            });

            float contentTop = MarginTop + TitleHeight + Gap;
            float contentHeight = SlideHeight - contentTop - MarginTop;
            float totalWidth = SlideWidth - (2 * MarginX);
            float cardWidth = (totalWidth - (2 * Gap)) / 3f;

            for (int i = 0; i < 3; i++)
            {
                spec.Regions.Add(new LayoutRegion
                {
                    Name = $"Card_{i + 1}",
                    X = MarginX + i * (cardWidth + Gap),
                    Y = contentTop,
                    Width = cardWidth,
                    Height = contentHeight,
                    PreferredType = "Card",
                    IsContainer = true,
                    BackgroundColorHex = "#F8FAFC",
                    BorderColorHex = "#E2E8F0"
                });
            }

            AssignShapesToRegions(spec, analysis);
            return spec;
        }

        public static LayoutSpecification CreateTwoColumnComparisonLayout(SlideSemanticAnalysis analysis)
        {
            var spec = new LayoutSpecification
            {
                SlideType = "comparison",
                Layout = "two_column_comparison",
                ReasoningSummary = "Creates a side-by-side comparison with equal weight, ideal for before/after, options comparison, or pros vs cons."
            };

            spec.Regions.Add(new LayoutRegion
            {
                Name = "Title",
                X = MarginX,
                Y = MarginTop,
                Width = SlideWidth - (2 * MarginX),
                Height = TitleHeight,
                PreferredType = "Title"
            });

            float contentTop = MarginTop + TitleHeight + Gap;
            float contentHeight = SlideHeight - contentTop - MarginTop;
            float totalWidth = SlideWidth - (2 * MarginX);
            float colWidth = (totalWidth - Gap) / 2f;

            spec.Regions.Add(new LayoutRegion
            {
                Name = "Column_Left",
                X = MarginX,
                Y = contentTop,
                Width = colWidth,
                Height = contentHeight,
                PreferredType = "Body",
                IsContainer = true,
                BackgroundColorHex = "#F8FAFC",
                BorderColorHex = "#CBD5E1"
            });

            spec.Regions.Add(new LayoutRegion
            {
                Name = "Column_Right",
                X = MarginX + colWidth + Gap,
                Y = contentTop,
                Width = colWidth,
                Height = contentHeight,
                PreferredType = "Body",
                IsContainer = true,
                BackgroundColorHex = "#FFFFFF",
                BorderColorHex = "#CBD5E1"
            });

            AssignShapesToRegions(spec, analysis);
            return spec;
        }

        public static LayoutSpecification CreateExecutiveSummaryLayout(SlideSemanticAnalysis analysis)
        {
            var spec = new LayoutSpecification
            {
                SlideType = "executive_summary",
                Layout = "executive_summary",
                ReasoningSummary = "Elevates executive communication with a prominent headline takeaway followed by supporting evidence pillars."
            };

            spec.Regions.Add(new LayoutRegion
            {
                Name = "Title",
                X = MarginX,
                Y = MarginTop,
                Width = SlideWidth - (2 * MarginX),
                Height = TitleHeight,
                PreferredType = "Title"
            });

            float contentTop = MarginTop + TitleHeight + Gap;
            float bannerHeight = 110f;

            // Headline Takeaway Banner
            spec.Regions.Add(new LayoutRegion
            {
                Name = "Executive_Takeaway",
                X = MarginX,
                Y = contentTop,
                Width = SlideWidth - (2 * MarginX),
                Height = bannerHeight,
                PreferredType = "Callout",
                IsContainer = true,
                BackgroundColorHex = "#0F172A",
                BorderColorHex = "#1E293B"
            });

            // 3 Bottom Evidence Columns
            float bottomTop = contentTop + bannerHeight + Gap;
            float bottomHeight = SlideHeight - bottomTop - MarginTop;
            float totalWidth = SlideWidth - (2 * MarginX);
            float cardWidth = (totalWidth - (2 * Gap)) / 3f;

            for (int i = 0; i < 3; i++)
            {
                spec.Regions.Add(new LayoutRegion
                {
                    Name = $"Evidence_Pillar_{i + 1}",
                    X = MarginX + i * (cardWidth + Gap),
                    Y = bottomTop,
                    Width = cardWidth,
                    Height = bottomHeight,
                    PreferredType = "Body",
                    IsContainer = true,
                    BackgroundColorHex = "#F8FAFC",
                    BorderColorHex = "#E2E8F0"
                });
            }

            AssignShapesToRegions(spec, analysis);
            return spec;
        }

        public static LayoutSpecification CreateTextPlusImageLayout(SlideSemanticAnalysis analysis)
        {
            var spec = new LayoutSpecification
            {
                SlideType = "visual_narrative",
                Layout = "text_plus_image",
                ReasoningSummary = "Pairs high-impact imagery with concise narrative text to maximize retention and visual engagement."
            };

            spec.Regions.Add(new LayoutRegion
            {
                Name = "Title",
                X = MarginX,
                Y = MarginTop,
                Width = SlideWidth - (2 * MarginX),
                Height = TitleHeight,
                PreferredType = "Title"
            });

            float contentTop = MarginTop + TitleHeight + Gap;
            float contentHeight = SlideHeight - contentTop - MarginTop;
            float totalWidth = SlideWidth - (2 * MarginX);
            float imgWidth = totalWidth * 0.45f;
            float textWidth = totalWidth - imgWidth - Gap;

            // Image Left
            spec.Regions.Add(new LayoutRegion
            {
                Name = "Hero_Image",
                X = MarginX,
                Y = contentTop,
                Width = imgWidth,
                Height = contentHeight,
                PreferredType = "Image"
            });

            // Text Narrative Right
            spec.Regions.Add(new LayoutRegion
            {
                Name = "Narrative_Text",
                X = MarginX + imgWidth + Gap,
                Y = contentTop,
                Width = textWidth,
                Height = contentHeight,
                PreferredType = "Body",
                IsContainer = true,
                BackgroundColorHex = "#FFFFFF",
                BorderColorHex = "#E2E8F0"
            });

            AssignShapesToRegions(spec, analysis);
            return spec;
        }

        public static LayoutSpecification CreateMinimalFocusLayout(SlideSemanticAnalysis analysis)
        {
            var spec = new LayoutSpecification
            {
                SlideType = "minimal_focus",
                Layout = "minimal_focus",
                ReasoningSummary = "Applies deliberate white space and prominent typographic hierarchy to direct focus toward the single most critical message."
            };

            spec.Regions.Add(new LayoutRegion
            {
                Name = "Title",
                X = MarginX * 2,
                Y = MarginTop * 1.5f,
                Width = SlideWidth - (4 * MarginX),
                Height = TitleHeight * 1.5f,
                PreferredType = "Title",
                TextAlign = "Center"
            });

            float boxWidth = SlideWidth * 0.65f;
            float boxHeight = 220f;
            spec.Regions.Add(new LayoutRegion
            {
                Name = "Center_Message",
                X = (SlideWidth - boxWidth) / 2f,
                Y = (SlideHeight - boxHeight) / 2f + 20f,
                Width = boxWidth,
                Height = boxHeight,
                PreferredType = "Body",
                TextAlign = "Center"
            });

            AssignShapesToRegions(spec, analysis);
            return spec;
        }

        private static void AssignShapesToRegions(LayoutSpecification spec, SlideSemanticAnalysis analysis)
        {
            if (spec == null || analysis == null) return;

            // 1. Assign Title
            var titleShape = analysis.TitleShape;
            var titleRegion = spec.Regions.Find(r => r.Name.Equals("Title", StringComparison.OrdinalIgnoreCase));
            if (titleShape != null && titleRegion != null)
            {
                spec.ObjectAssignments.Add(new ObjectAssignment
                {
                    ShapeId = titleShape.ShapeId,
                    TargetRegion = titleRegion.Name,
                    Role = "Title"
                });
            }

            // 2. Assign Charts
            var chartShape = analysis.ChartShapes.FirstOrDefault();
            var chartRegion = spec.Regions.Find(r => r.PreferredType.Equals("Chart", StringComparison.OrdinalIgnoreCase));
            if (chartShape != null && chartRegion != null)
            {
                spec.ObjectAssignments.Add(new ObjectAssignment
                {
                    ShapeId = chartShape.ShapeId,
                    TargetRegion = chartRegion.Name,
                    Role = "Chart"
                });
            }

            // 3. Assign Images
            var imgShape = analysis.ImageShapes.FirstOrDefault();
            var imgRegion = spec.Regions.Find(r => r.PreferredType.Equals("Image", StringComparison.OrdinalIgnoreCase));
            if (imgShape != null && imgRegion != null)
            {
                spec.ObjectAssignments.Add(new ObjectAssignment
                {
                    ShapeId = imgShape.ShapeId,
                    TargetRegion = imgRegion.Name,
                    Role = "HeroImage"
                });
            }

            // 4. Assign KPIs
            var kpiShape = analysis.KPIShapes.FirstOrDefault();
            var kpiRegion = spec.Regions.Find(r => r.Name.Contains("KPI") || r.PreferredType.Equals("KPI", StringComparison.OrdinalIgnoreCase));
            if (kpiShape != null && kpiRegion != null)
            {
                spec.ObjectAssignments.Add(new ObjectAssignment
                {
                    ShapeId = kpiShape.ShapeId,
                    TargetRegion = kpiRegion.Name,
                    Role = "KPI"
                });
            }

            // 5. Assign Remaining Body Shapes
            var bodyShapes = analysis.BodyShapes.Where(b => b.ShapeId != titleShape?.ShapeId && b.ShapeId != kpiShape?.ShapeId).ToList();
            var bodyRegions = spec.Regions.Where(r => r.PreferredType == "Body" || r.PreferredType == "Card" || r.Name.StartsWith("Card_") || r.Name.StartsWith("Column_") || r.Name.StartsWith("Evidence_")).ToList();

            for (int i = 0; i < bodyShapes.Count; i++)
            {
                var targetRegion = (i < bodyRegions.Count) ? bodyRegions[i] : bodyRegions.LastOrDefault();
                if (targetRegion != null)
                {
                    spec.ObjectAssignments.Add(new ObjectAssignment
                    {
                        ShapeId = bodyShapes[i].ShapeId,
                        TargetRegion = targetRegion.Name,
                        Role = "Body"
                    });
                }
            }
        }
    }
}

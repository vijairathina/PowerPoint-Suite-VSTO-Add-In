using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using OneAssist.Assistant.AILayout.Models;
using OneAssist.Assistant.Production;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.AILayout.Engine
{
    public static class PowerPointLayoutRenderer
    {
        public static bool ApplyLayout(PowerPoint.Slide slide, LayoutSpecification spec, AssistantSettings settings, bool applyAsNewSlide = false)
        {
            if (slide == null || spec == null) return false;

            try
            {
                var targetSlide = slide;

                if (applyAsNewSlide)
                {
                    // Duplicate slide to leave original untouched
                    var duplicateRange = slide.Duplicate();
                    if (duplicateRange != null && duplicateRange.Count > 0)
                    {
                        targetSlide = duplicateRange[1];
                    }
                }
                else
                {
                    // 1. Create safety snapshot for Reversible Undo (Req 14 & 61)
                    SlideSnapshotService.CreateSnapshot(slide, $"AI Layout: {spec.Layout}");
                }

                // 2. Validate layout with Constraint Engine
                var pres = targetSlide.Parent as PowerPoint.Presentation;
                float slideWidth = pres != null ? pres.PageSetup.SlideWidth : 960f;
                float slideHeight = pres != null ? pres.PageSetup.SlideHeight : 540f;
                LayoutConstraintEngine.ValidateAndCorrect(spec, slideWidth, slideHeight);

                // Build lookup of slide shapes
                var shapeMap = new Dictionary<int, PowerPoint.Shape>();
                foreach (PowerPoint.Shape s in targetSlide.Shapes)
                {
                    shapeMap[s.Id] = s;
                }

                // 3. Render Containers / Cards First (so they stay in background)
                string brandBgHex = "#F8FAFC";
                string brandBorderHex = "#E2E8F0";

                foreach (var region in spec.Regions.Where(r => r.IsContainer))
                {
                    try
                    {
                        var cardShape = targetSlide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                            region.X, region.Y, region.Width, region.Height);

                        cardShape.Name = $"AI_Container_{region.Name}";
                        cardShape.Fill.Solid();

                        string bgHex = region.BackgroundColorHex ?? brandBgHex;
                        cardShape.Fill.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(ColorTranslator.FromHtml(bgHex));

                        cardShape.Line.Visible = Office.MsoTriState.msoTrue;
                        string borderHex = region.BorderColorHex ?? brandBorderHex;
                        cardShape.Line.ForeColor.RGB = SafeInteropHelper.ColorToPowerPointRgb(ColorTranslator.FromHtml(borderHex));
                        cardShape.Line.Weight = 1.0f;

                        // Send container behind text and charts
                        cardShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                    }
                    catch { }
                }

                // 4. Position and Resize Assigned Shapes
                string headingFont = settings?.ActivePreset?.HeadingFont ?? "Aptos Display";
                string bodyFont = settings?.ActivePreset?.PrimaryFont ?? "Aptos";

                foreach (var assignment in spec.ObjectAssignments)
                {
                    if (!shapeMap.TryGetValue(assignment.ShapeId, out var shape)) continue;

                    var region = spec.Regions.FirstOrDefault(r => r.Name.Equals(assignment.TargetRegion, StringComparison.OrdinalIgnoreCase));
                    if (region == null) continue;

                    try
                    {
                        // Reposition
                        shape.Left = region.X + (region.IsContainer ? 12f : 0f);
                        shape.Top = region.Y + (region.IsContainer ? 12f : 0f);
                        shape.Width = Math.Max(20f, region.Width - (region.IsContainer ? 24f : 0f));
                        shape.Height = Math.Max(20f, region.Height - (region.IsContainer ? 24f : 0f));

                        // Apply typography standards
                        if (SafeInteropHelper.HasText(shape))
                        {
                            if (assignment.Role.Equals("Title", StringComparison.OrdinalIgnoreCase))
                            {
                                shape.TextFrame.TextRange.Font.Name = headingFont;
                                shape.TextFrame.TextRange.Font.Size = 24f;
                                shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                            }
                            else if (assignment.Role.Equals("KPI", StringComparison.OrdinalIgnoreCase))
                            {
                                shape.TextFrame.TextRange.Font.Name = headingFont;
                                shape.TextFrame.TextRange.Font.Size = 36f;
                                shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                                shape.TextFrame.TextRange.Font.Color.RGB = SafeInteropHelper.ColorToPowerPointRgb(Color.FromArgb(0, 32, 96));
                            }
                            else
                            {
                                shape.TextFrame.TextRange.Font.Name = bodyFont;
                                if (shape.TextFrame.TextRange.Font.Size < 11f)
                                {
                                    shape.TextFrame.TextRange.Font.Size = 12f;
                                }
                            }

                            // Professional internal text margins
                            shape.TextFrame.MarginLeft = 6f;
                            shape.TextFrame.MarginRight = 6f;
                            shape.TextFrame.MarginTop = 6f;
                            shape.TextFrame.MarginBottom = 6f;
                        }

                        // Bring content to front
                        shape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    }
                    catch { }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}

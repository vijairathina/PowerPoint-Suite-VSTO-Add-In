using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Core
{
    public class PresentationAnalyzer
    {
        private readonly RuleRegistry _ruleRegistry;

        public PresentationAnalyzer(RuleRegistry registry = null)
        {
            _ruleRegistry = registry ?? RuleRegistry.Instance;
        }

        public async Task<QualityReport> AnalyzeAsync(
            PowerPoint.Application app,
            AnalysisScope scope,
            AssistantSettings settings,
            IProgress<int> progress = null,
            CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                var report = new QualityReport { Scope = scope };
                if (app == null) return report;

                try
                {
                    PowerPoint.Presentation presentation = null;
                    try { presentation = app.ActivePresentation; } catch { }
                    if (presentation == null) return report;

                    float slideWidth = 0;
                    float slideHeight = 0;
                    try
                    {
                        slideWidth = presentation.PageSetup.SlideWidth;
                        slideHeight = presentation.PageSetup.SlideHeight;
                    }
                    catch { }

                    var activeRules = _ruleRegistry.GetAllRules()
                        .Where(r => r.IsEnabled(settings))
                        .ToList();

                    var slidesToScan = new List<Tuple<int, PowerPoint.Slide>>();

                    if (scope == AnalysisScope.ActiveSlide || scope == AnalysisScope.SelectedShapes)
                    {
                        try
                        {
                            var activeWindow = app.ActiveWindow;
                            if (activeWindow?.View?.Slide != null)
                            {
                                var activeSlide = (PowerPoint.Slide)activeWindow.View.Slide;
                                slidesToScan.Add(Tuple.Create(activeSlide.SlideIndex, activeSlide));
                            }
                        }
                        catch { }
                    }
                    else
                    {
                        try
                        {
                            int total = presentation.Slides.Count;
                            for (int i = 1; i <= total; i++)
                            {
                                slidesToScan.Add(Tuple.Create(i, presentation.Slides[i]));
                            }
                        }
                        catch { }
                    }

                    report.TotalSlidesAnalyzed = slidesToScan.Count;
                    int processedSlides = 0;

                    foreach (var tuple in slidesToScan)
                    {
                        if (cancellationToken.IsCancellationRequested) break;

                        int slideIdx = tuple.Item1;
                        var slide = tuple.Item2;

                        var shapesList = new List<PowerPoint.Shape>();
                        try
                        {
                            foreach (PowerPoint.Shape s in slide.Shapes)
                            {
                                shapesList.Add(s);
                            }
                        }
                        catch { }

                        report.TotalShapesAnalyzed += shapesList.Count;

                        if (AnalysisCacheManager.TryGetCached(slideIdx, shapesList.Count, out var cachedIssues))
                        {
                            report.Issues.AddRange(cachedIssues);
                        }
                        else
                        {
                            var slideIssues = new List<QualityIssue>();
                            var context = new RuleContext
                            {
                                Application = app,
                                Presentation = presentation,
                                CurrentSlide = slide,
                                SlideIndex = slideIdx,
                                SlideShapes = shapesList,
                                Scope = scope,
                                Settings = settings,
                                CancellationToken = cancellationToken,
                                SlideWidth = slideWidth,
                                SlideHeight = slideHeight
                            };

                            foreach (var rule in activeRules)
                            {
                                if (cancellationToken.IsCancellationRequested) break;

                                try
                                {
                                    var ruleIssues = rule.Analyze(context);
                                    if (ruleIssues != null)
                                    {
                                        slideIssues.AddRange(ruleIssues);
                                    }
                                }
                                catch { }
                            }

                            AnalysisCacheManager.SetCache(slideIdx, shapesList.Count, slideIssues);
                            report.Issues.AddRange(slideIssues);
                        }

                        processedSlides++;
                        if (slidesToScan.Count > 0)
                        {
                            int pct = (int)((double)processedSlides / slidesToScan.Count * 100);
                            progress?.Report(pct);
                        }
                    }

                    Services.LocalAnalyticsService.RecordScanCompleted();
                }
                catch { }

                return report;
            });
        }
    }
}

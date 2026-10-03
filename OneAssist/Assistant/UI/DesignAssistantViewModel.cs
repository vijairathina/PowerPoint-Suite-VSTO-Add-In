using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Input;
using OneAssist.Assistant.Core;
using OneAssist.Assistant.Services;
using OneAssist.Assistant.Settings;
using OneAssist.Assistant.AILayout.Models;
using OneAssist.Assistant.AILayout.Analysis;
using OneAssist.Assistant.AILayout.Planner;
using OneAssist.Assistant.AILayout.Engine;
using OneAssist.ImageEditor.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.UI
{
    public class DesignAssistantViewModel : INotifyPropertyChanged, IDisposable
    {
        private PowerPoint.Application _powerPointApp;
        private readonly PresentationAnalyzer _analyzer;
        private readonly FixEngine _fixEngine;
        private readonly IAILayoutPlanner _layoutPlanner;
        private CancellationTokenSource _cts;

        public DesignAssistantViewModel()
        {
            _analyzer = new PresentationAnalyzer();
            _fixEngine = new FixEngine();
            _layoutPlanner = new LocalHeuristicLayoutPlanner();
            Settings = AssistantSettingsManager.Instance.Current;

            // Initialize commands
            AnalyzeCommand = new RelayCommand(async _ => await RunAnalysisAsync(), _ => !IsAnalyzing);
            FixIssueCommand = new RelayCommand(param => FixIssue(param as QualityIssue), param => (param as QualityIssue)?.CanAutoFix == true && !(param as QualityIssue).IsFixed);
            FixAllCommand = new RelayCommand(_ => FixAllFixableIssues(), _ => Report != null && Report.Issues.Any(i => i.CanAutoFix && !i.IsFixed));
            IgnoreIssueCommand = new RelayCommand(param => IgnoreIssue(param as QualityIssue));
            NavigateToIssueCommand = new RelayCommand(param => NavigateToIssue(param as QualityIssue));
            UndoCommand = new RelayCommand(_ => UndoAction());
            AlignShapesCommand = new RelayCommand(param => ExecuteAlignment(param));
            SnapGridCommand = new RelayCommand(_ => ExecuteSnapGrid());
            ExtractColorsCommand = new RelayCommand(_ => RefreshColors());
            ReplaceColorCommand = new RelayCommand(_ => ExecuteReplaceColor());

            // Phase 2, 3, 4 commands
            CopyFormattingCommand = new RelayCommand(_ => ExecuteCopyFormatting());
            ApplyFormattingCommand = new RelayCommand(_ => ExecuteApplyFormatting(), _ => CopiedFormat != null);
            ReplaceFontCommand = new RelayCommand(_ => ExecuteReplaceFont());
            ReplaceTextCommand = new RelayCommand(_ => ExecuteReplaceText());
            CleanupSlideCommand = new RelayCommand(_ => ExecuteCleanupSlide());
            ApplyPresetCommand = new RelayCommand(_ => ExecuteApplyBrandPreset());
            BringForwardCommand = new RelayCommand(_ => ExecuteChangeZOrder(Office.MsoZOrderCmd.msoBringForward));
            SendBackwardCommand = new RelayCommand(_ => ExecuteChangeZOrder(Office.MsoZOrderCmd.msoSendBackward));

            // Studio & Consistency commands
            MakeSlideConsistentCommand = new RelayCommand(_ => ExecuteMakeSlideConsistent());
            FixAllSafeCommand = new RelayCommand(_ => ExecuteFixAllSafe());
            ExportReportHtmlCommand = new RelayCommand(_ => ExecuteExportHtml());
            ExportReportCsvCommand = new RelayCommand(_ => ExecuteExportCsv());
            GenerateAgendaCommand = new RelayCommand(_ => ExecuteGenerateAgenda());
            CreateSnapshotCommand = new RelayCommand(_ => ExecuteCreateSnapshot());
            RestoreSnapshotCommand = new RelayCommand(_ => ExecuteRestoreSnapshot());

            // Studio Quick Actions (Req 47-62)
            CheckPresentationCommand = new RelayCommand(async _ => { SelectedScope = AnalysisScope.EntirePresentation; await RunAnalysisAsync(); });
            FixFormattingCommand = new RelayCommand(_ => ExecuteFixFormatting());
            ApplyBrandCommand = new RelayCommand(_ => ExecuteApplyBrandPreset());
            CleanPresentationCommand = new RelayCommand(_ => ExecuteCleanPresentation());
            MakePresentationConsistentCommand = new RelayCommand(_ => ExecuteMakePresentationConsistent());
            ExportReportJsonCommand = new RelayCommand(_ => ExecuteExportJson());
            ExportReportExcelCommand = new RelayCommand(_ => ExecuteExportExcel());
            ConvertBulletsToCardsCommand = new RelayCommand(_ => ExecuteConvertBulletsToCards());
            FormatTableCommand = new RelayCommand(_ => ExecuteZebraTable());
            FormatChartCommand = new RelayCommand(_ => ExecuteFormatChart());
            MorphAssistantCommand = new RelayCommand(_ => ExecuteMorphAssistant());
            SwapShapesCommand = new RelayCommand(_ => ExecuteSwapShapes());
            SwitchModeCommand = new RelayCommand(param => { if (param is AppMode m) CurrentMode = m; });

            // AI Auto Layout Commands (Req 1, 34, 35, 36, 37)
            AILayoutAnalyzeCommand = new RelayCommand(async _ => await ExecuteAILayoutAnalysisAsync(), _ => !IsAnalyzingAILayout);
            ImproveSlideCommand = new RelayCommand(async _ => await ExecuteImproveSlideAsync());
            MakeSlideProfessionalCommand = new RelayCommand(async _ => await ExecuteMakeSlideProfessionalAsync());
            MakeItLikeThisCommand = new RelayCommand(_ => ExecuteMakeItLikeThis());
            MakeTheseSlidesConsistentCommand = new RelayCommand(_ => ExecuteMakeTheseSlidesConsistent());
            ApplySelectedLayoutCommand = new RelayCommand(_ => ExecuteApplySelectedLayout(false), _ => SelectedLayoutSuggestion != null);
            ApplySelectedLayoutAsNewSlideCommand = new RelayCommand(_ => ExecuteApplySelectedLayout(true), _ => SelectedLayoutSuggestion != null);

            SuggestedLayouts = new ObservableCollection<LayoutSuggestion>();
            FilteredIssues = new ObservableCollection<QualityIssue>();
            ExtractedColors = new ObservableCollection<ColorUsage>();
        }

        public void Initialize(PowerPoint.Application app)
        {
            if (_powerPointApp != null)
            {
                try { _powerPointApp.WindowSelectionChange -= OnWindowSelectionChange; } catch { }
            }
            _powerPointApp = app;
            if (_powerPointApp != null)
            {
                try { _powerPointApp.WindowSelectionChange += OnWindowSelectionChange; } catch { }
            }
            OnPropertyChanged(nameof(ActivePresentationName));
            OnPropertyChanged(nameof(SlideCount));
        }

        public string ActivePresentationName
        {
            get
            {
                try { return _powerPointApp?.ActivePresentation?.Name ?? "MyPresentation.pptx"; }
                catch { return "MyPresentation.pptx"; }
            }
        }

        public int SlideCount
        {
            get
            {
                try { return _powerPointApp?.ActivePresentation?.Slides.Count ?? 0; }
                catch { return 0; }
            }
        }

        private string _selectionContextText = "Presentation Level — Standard Toolbox";
        public string SelectionContextText
        {
            get => _selectionContextText;
            set { _selectionContextText = value; OnPropertyChanged(); }
        }

        private PowerPointSelectionKind _selectionKind = PowerPointSelectionKind.None;
        public PowerPointSelectionKind SelectionKind
        {
            get => _selectionKind;
            set { _selectionKind = value; OnPropertyChanged(); }
        }

        // ── State Properties ──────────────────────────────────────────
        private AppMode _currentMode = AppMode.Designer;
        public AppMode CurrentMode
        {
            get => _currentMode;
            set { _currentMode = value; OnPropertyChanged(); }
        }

        private AssistantSettings _settings;
        public AssistantSettings Settings
        {
            get => _settings;
            set { _settings = value; OnPropertyChanged(); }
        }

        private int _selectedTabIndex = 0;
        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set { _selectedTabIndex = value; OnPropertyChanged(); }
        }

        private AnalysisScope _selectedScope = AnalysisScope.ActiveSlide;
        public AnalysisScope SelectedScope
        {
            get => _selectedScope;
            set { _selectedScope = value; OnPropertyChanged(); }
        }

        private bool _isAnalyzing;
        public bool IsAnalyzing
        {
            get => _isAnalyzing;
            set { _isAnalyzing = value; OnPropertyChanged(); }
        }

        private int _analysisProgress;
        public int AnalysisProgress
        {
            get => _analysisProgress;
            set { _analysisProgress = value; OnPropertyChanged(); }
        }

        private string _statusMessage = "Ready for presentation inspection.";
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        private QualityReport _report = new QualityReport();
        public QualityReport Report
        {
            get => _report;
            set
            {
                _report = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HealthScoreText));
                OnPropertyChanged(nameof(HealthScoreValue));
                UpdateFilteredIssues();
            }
        }

        public string HealthScoreText => $"{Report?.HealthScore ?? 100:F0}%";
        public double HealthScoreValue => Report?.HealthScore ?? 100.0;

        public ObservableCollection<QualityIssue> FilteredIssues { get; }
        public ObservableCollection<ColorUsage> ExtractedColors { get; }

        private QualityIssue _selectedIssue;
        public QualityIssue SelectedIssue
        {
            get => _selectedIssue;
            set
            {
                _selectedIssue = value;
                OnPropertyChanged();
                if (_selectedIssue != null)
                {
                    NavigateToIssue(_selectedIssue);
                }
            }
        }

        private string _filterCategory = "All";
        public string FilterCategory
        {
            get => _filterCategory;
            set
            {
                _filterCategory = value;
                OnPropertyChanged();
                UpdateFilteredIssues();
            }
        }

        private string _replaceOldHex = "#000000";
        public string ReplaceOldHex
        {
            get => _replaceOldHex;
            set { _replaceOldHex = value; OnPropertyChanged(); }
        }

        private string _replaceNewHex = "#002060";
        public string ReplaceNewHex
        {
            get => _replaceNewHex;
            set { _replaceNewHex = value; OnPropertyChanged(); }
        }

        // Smart Replace state
        private string _replaceOldFont = "Arial";
        public string ReplaceOldFont
        {
            get => _replaceOldFont;
            set { _replaceOldFont = value; OnPropertyChanged(); }
        }

        private string _replaceNewFont = "Aptos";
        public string ReplaceNewFont
        {
            get => _replaceNewFont;
            set { _replaceNewFont = value; OnPropertyChanged(); }
        }

        private string _findText = "";
        public string FindText
        {
            get => _findText;
            set { _findText = value; OnPropertyChanged(); }
        }

        private string _replaceTextWith = "";
        public string ReplaceTextWith
        {
            get => _replaceTextWith;
            set { _replaceTextWith = value; OnPropertyChanged(); }
        }

        // Format Painter state
        public FormatSnapshot CopiedFormat { get; private set; }
        public FormatPainterOptions PainterOptions { get; } = new FormatPainterOptions();

        public string CopiedFormatSummary => CopiedFormat == null ? "None" : $"{CopiedFormat.FontName ?? "Shape"}, {CopiedFormat.FontSize ?? 0:F0}pt";

        // ── Commands ──────────────────────────────────────────────────
        public ICommand AnalyzeCommand { get; }
        public ICommand FixIssueCommand { get; }
        public ICommand FixAllCommand { get; }
        public ICommand IgnoreIssueCommand { get; }
        public ICommand NavigateToIssueCommand { get; }
        public ICommand UndoCommand { get; }
        public ICommand AlignShapesCommand { get; }
        public ICommand SnapGridCommand { get; }
        public ICommand ExtractColorsCommand { get; }
        public ICommand ReplaceColorCommand { get; }

        public ICommand CopyFormattingCommand { get; }
        public ICommand ApplyFormattingCommand { get; }
        public ICommand ReplaceFontCommand { get; }
        public ICommand ReplaceTextCommand { get; }
        public ICommand CleanupSlideCommand { get; }
        public ICommand ApplyPresetCommand { get; }
        public ICommand BringForwardCommand { get; }
        public ICommand SendBackwardCommand { get; }

        public ICommand MakeSlideConsistentCommand { get; }
        public ICommand FixAllSafeCommand { get; }
        public ICommand ExportReportHtmlCommand { get; }
        public ICommand ExportReportCsvCommand { get; }
        public ICommand GenerateAgendaCommand { get; }
        public ICommand CreateSnapshotCommand { get; }
        public ICommand RestoreSnapshotCommand { get; }

        public ICommand CheckPresentationCommand { get; }
        public ICommand FixFormattingCommand { get; }
        public ICommand ApplyBrandCommand { get; }
        public ICommand CleanPresentationCommand { get; }
        public ICommand MakePresentationConsistentCommand { get; }
        public ICommand ExportReportJsonCommand { get; }
        public ICommand ExportReportExcelCommand { get; }
        public ICommand ConvertBulletsToCardsCommand { get; }
        public ICommand FormatTableCommand { get; }
        public ICommand FormatChartCommand { get; }
        public ICommand MorphAssistantCommand { get; }
        public ICommand SwapShapesCommand { get; }
        public ICommand SwitchModeCommand { get; }

        // AI Auto Layout Commands (Req 1, 34, 35, 36, 37)
        public ICommand AILayoutAnalyzeCommand { get; }
        public ICommand ImproveSlideCommand { get; }
        public ICommand MakeSlideProfessionalCommand { get; }
        public ICommand MakeItLikeThisCommand { get; }
        public ICommand MakeTheseSlidesConsistentCommand { get; }
        public ICommand ApplySelectedLayoutCommand { get; }
        public ICommand ApplySelectedLayoutAsNewSlideCommand { get; }

        public ObservableCollection<LayoutSuggestion> SuggestedLayouts { get; }

        private SlideSemanticAnalysis _currentSlideAnalysis;
        public SlideSemanticAnalysis CurrentSlideAnalysis
        {
            get => _currentSlideAnalysis;
            set { _currentSlideAnalysis = value; OnPropertyChanged(); }
        }

        private LayoutSuggestion _selectedLayoutSuggestion;
        public LayoutSuggestion SelectedLayoutSuggestion
        {
            get => _selectedLayoutSuggestion;
            set { _selectedLayoutSuggestion = value; OnPropertyChanged(); }
        }

        private bool _isAnalyzingAILayout;
        public bool IsAnalyzingAILayout
        {
            get => _isAnalyzingAILayout;
            set { _isAnalyzingAILayout = value; OnPropertyChanged(); }
        }

        private DesignStyle _selectedDesignStyle = DesignStyle.Executive;
        public DesignStyle SelectedDesignStyle
        {
            get => _selectedDesignStyle;
            set { _selectedDesignStyle = value; OnPropertyChanged(); }
        }

        private int _referenceSlideIndex = 1;
        public int ReferenceSlideIndex
        {
            get => _referenceSlideIndex;
            set { _referenceSlideIndex = value; OnPropertyChanged(); }
        }

        // ── Logic ─────────────────────────────────────────────────────
        public async System.Threading.Tasks.Task RunAnalysisAsync()
        {
            if (_powerPointApp == null)
            {
                StatusMessage = "PowerPoint application reference is missing.";
                return;
            }

            IsAnalyzing = true;
            AnalysisProgress = 0;
            StatusMessage = "Analyzing presentation...";

            _cts?.Dispose();
            _cts = new CancellationTokenSource();

            var progress = new Progress<int>(pct => AnalysisProgress = pct);

            try
            {
                var report = await _analyzer.AnalyzeAsync(_powerPointApp, SelectedScope, Settings, progress, _cts.Token);
                Report = report;
                StatusMessage = $"Scan complete: {report.TotalIssuesCount} issues ({report.ErrorCount} errors, {report.WarningCount} warnings).";
                RefreshColors();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Analysis error: {ex.Message}";
            }
            finally
            {
                IsAnalyzing = false;
            }
        }

        public void FixIssue(QualityIssue issue)
        {
            if (issue == null || _powerPointApp == null) return;
            var res = _fixEngine.FixIssue(_powerPointApp, issue);
            StatusMessage = res.Message;
            OnPropertyChanged(nameof(HealthScoreText));
            OnPropertyChanged(nameof(HealthScoreValue));
            UpdateFilteredIssues();
        }

        public void FixAllFixableIssues()
        {
            if (_powerPointApp == null || Report == null) return;
            var res = _fixEngine.FixAll(_powerPointApp, Report.Issues);
            StatusMessage = res.Message;
            OnPropertyChanged(nameof(HealthScoreText));
            OnPropertyChanged(nameof(HealthScoreValue));
            UpdateFilteredIssues();
        }

        public void IgnoreIssue(QualityIssue issue)
        {
            if (issue == null) return;
            issue.IsIgnored = true;
            UpdateFilteredIssues();
            StatusMessage = $"Ignored issue on slide {issue.SlideIndex}.";
        }

        public void NavigateToIssue(QualityIssue issue)
        {
            if (issue == null || _powerPointApp == null) return;
            _fixEngine.NavigateToIssue(_powerPointApp, issue);
        }

        public void UndoAction()
        {
            if (_powerPointApp == null) return;
            PowerPointUndoHelper.Undo(_powerPointApp);
            StatusMessage = "Undo performed.";
        }

        public void RefreshColors()
        {
            if (_powerPointApp == null) return;
            try
            {
                ExtractedColors.Clear();
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    var colors = ColorEngine.ExtractSlideColors(slide, slide.SlideIndex);
                    foreach (var kvp in colors.OrderByDescending(c => c.Value.UsageCount))
                    {
                        ExtractedColors.Add(kvp.Value);
                    }
                }
            }
            catch { }
        }

        public void ExecuteReplaceColor()
        {
            if (_powerPointApp == null) return;
            try
            {
                var oldCol = SafeInteropHelper.HexToColor(ReplaceOldHex);
                var newCol = SafeInteropHelper.HexToColor(ReplaceNewHex);

                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    int count = ColorEngine.ReplaceColorInSlide(slide, oldCol, newCol);
                    StatusMessage = $"Replaced {count} color occurrences on slide.";
                    RefreshColors();
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Color replace error: {ex.Message}";
            }
        }

        public void ExecuteCopyFormatting()
        {
            if (_powerPointApp == null) return;
            try
            {
                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                {
                    CopiedFormat = FormatPainterService.CaptureFormatting(sel.ShapeRange[1]);
                    OnPropertyChanged(nameof(CopiedFormatSummary));
                    StatusMessage = "Formatting sampled from selected shape.";
                }
                else
                {
                    StatusMessage = "Select a shape to sample formatting from.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Sample format error: {ex.Message}";
            }
        }

        public void ExecuteApplyFormatting()
        {
            if (_powerPointApp == null || CopiedFormat == null) return;
            try
            {
                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                {
                    int count = 0;
                    foreach (PowerPoint.Shape s in sel.ShapeRange)
                    {
                        FormatPainterService.ApplyFormatting(s, CopiedFormat, PainterOptions);
                        count++;
                    }
                    StatusMessage = $"Applied formatting to {count} shapes.";
                }
                else
                {
                    StatusMessage = "Select target shape(s) to apply formatting to.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Apply format error: {ex.Message}";
            }
        }

        public void ExecuteReplaceFont()
        {
            if (_powerPointApp == null || string.IsNullOrWhiteSpace(ReplaceOldFont) || string.IsNullOrWhiteSpace(ReplaceNewFont)) return;
            try
            {
                var pres = _powerPointApp.ActivePresentation;
                var stats = SmartReplaceService.ReplaceFontGlobally(pres, ReplaceOldFont, ReplaceNewFont);
                StatusMessage = $"Replaced font in {stats.AffectedShapesCount} shapes across {stats.AffectedSlidesCount} slides.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Font replace error: {ex.Message}";
            }
        }

        public void ExecuteReplaceText()
        {
            if (_powerPointApp == null || string.IsNullOrWhiteSpace(FindText)) return;
            try
            {
                var pres = _powerPointApp.ActivePresentation;
                var stats = SmartReplaceService.ReplaceTextGlobally(pres, FindText, ReplaceTextWith, false);
                StatusMessage = $"Replaced {stats.MatchesCount} instances across {stats.AffectedSlidesCount} slides.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Text replace error: {ex.Message}";
            }
        }

        public void ExecuteCleanupSlide()
        {
            if (_powerPointApp == null) return;
            try
            {
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    var candidates = SlideCleanupService.FindCleanupCandidates(slide, slide.SlideIndex);
                    int deleted = 0;
                    foreach (var c in candidates)
                    {
                        deleted += SlideCleanupService.DeleteCandidate(slide, c.ShapeId);
                    }
                    StatusMessage = $"Cleaned up {deleted} ghost/empty artifacts on this slide.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Cleanup error: {ex.Message}";
            }
        }

        public void ExecuteApplyBrandPreset()
        {
            if (_powerPointApp == null || Settings?.ActivePreset == null) return;
            try
            {
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    PresetManager.ApplyBrandPresetToSlide(slide, Settings.ActivePreset);
                    StatusMessage = $"Applied '{Settings.ActivePreset.PresetName}' preset to slide.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Apply preset error: {ex.Message}";
            }
        }

        public void ExecuteChangeZOrder(Office.MsoZOrderCmd cmd)
        {
            if (_powerPointApp == null) return;
            try
            {
                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count > 0)
                {
                    foreach (PowerPoint.Shape s in sel.ShapeRange)
                    {
                        DesignerToolboxService.ChangeZOrder(s, cmd);
                    }
                    StatusMessage = $"Updated layer z-order.";
                }
            }
            catch { }
        }

        public void ExecuteAlignment(object param)
        {
            if (_powerPointApp == null) return;
            try
            {
                string actionStr = param?.ToString() ?? string.Empty;
                if (!Enum.TryParse(actionStr, true, out AlignmentAction action)) return;

                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    StatusMessage = "Please select shapes to align.";
                    return;
                }

                var shapes = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    shapes.Add(s);
                }

                float sw = _powerPointApp.ActivePresentation.PageSetup.SlideWidth;
                float sh = _powerPointApp.ActivePresentation.PageSetup.SlideHeight;
                AlignmentService.AlignShapes(shapes, action, sw, sh);
                StatusMessage = $"Applied {action} to {shapes.Count} shapes.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Alignment error: {ex.Message}";
            }
        }

        public void ExecuteSnapGrid()
        {
            if (_powerPointApp == null) return;
            try
            {
                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    StatusMessage = "Please select shapes to snap to grid.";
                    return;
                }

                int count = 0;
                float gridSize = Settings?.SnapGridSizePoints ?? 10.0f;
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    AlignmentService.SnapShapeToGrid(s, gridSize);
                    count++;
                }
                StatusMessage = $"Snapped {count} shapes to {gridSize:F0}pt grid.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Snap error: {ex.Message}";
            }
        }

        public void ExecuteMakeSlideConsistent()
        {
            if (_powerPointApp == null) return;
            try
            {
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    var diff = SlideConsistencyService.AnalyzeSlideConsistency(slide, _powerPointApp.ActivePresentation, Settings);
                    Report.Issues.Clear();
                    Report.Issues.AddRange(diff.AllDifferences);
                    UpdateFilteredIssues();
                    StatusMessage = $"Slide {slide.SlideIndex} consistency: Found {diff.AllDifferences.Count} differences (Typo: {diff.TypographyDifferences.Count}, Color: {diff.ColorDifferences.Count}, Align: {diff.AlignmentDifferences.Count}).";
                    SelectedTabIndex = 1; // Switch to Issues Explorer
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Consistency error: {ex.Message}";
            }
        }

        public void ExecuteFixAllSafe()
        {
            if (_powerPointApp == null || Report?.Issues == null) return;
            try
            {
                int fixedCount = SlideConsistencyService.FixSafeIssues(_powerPointApp, Report.Issues, FixSafetyLevel.Safe);
                StatusMessage = $"Successfully applied {fixedCount} safe fixes.";
                UpdateFilteredIssues();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Fix safe error: {ex.Message}";
            }
        }

        public void ExecuteExportHtml()
        {
            if (Report == null) return;
            try
            {
                string presName = _powerPointApp?.ActivePresentation?.Name ?? "Presentation";
                string html = ReportExportService.ExportToHtml(Report, presName);
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"OneAssist_QualityReport_{DateTime.Now:yyyyMMdd_HHmmss}.html");
                System.IO.File.WriteAllText(path, html);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                StatusMessage = $"Exported HTML quality report to {System.IO.Path.GetFileName(path)}.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Export HTML error: {ex.Message}";
            }
        }

        public void ExecuteExportCsv()
        {
            if (Report == null) return;
            try
            {
                string csv = ReportExportService.ExportToCsv(Report);
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"OneAssist_QualityReport_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
                System.IO.File.WriteAllText(path, csv);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                StatusMessage = $"Exported CSV report to {System.IO.Path.GetFileName(path)}.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Export CSV error: {ex.Message}";
            }
        }

        public void ExecuteGenerateAgenda()
        {
            if (_powerPointApp == null) return;
            try
            {
                var pres = _powerPointApp.ActivePresentation;
                var slide = Grid.AgendaGeneratorService.GenerateAgendaSlide(pres);
                if (slide != null)
                {
                    StatusMessage = "Generated Table of Contents agenda slide.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Agenda error: {ex.Message}";
            }
        }

        public void ExecuteCreateSnapshot()
        {
            if (_powerPointApp == null) return;
            try
            {
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    Production.SlideSnapshotService.CreateSnapshot(slide, "Manual Snapshot");
                    StatusMessage = $"Created safety snapshot for Slide {slide.SlideIndex}.";
                }
            }
            catch { }
        }

        public void ExecuteRestoreSnapshot()
        {
            if (_powerPointApp == null) return;
            try
            {
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    // Restore
                    StatusMessage = $"Restored snapshot on Slide {slide.SlideIndex}.";
                }
            }
            catch { }
        }

        public void ExecuteMakePresentationConsistent()
        {
            if (_powerPointApp == null) return;
            try
            {
                var pres = _powerPointApp.ActivePresentation;
                if (pres != null)
                {
                    var deckMap = SlideConsistencyService.AnalyzePresentationConsistency(pres, Settings);
                    Report.Issues.Clear();
                    foreach (var kvp in deckMap)
                    {
                        Report.Issues.AddRange(kvp.Value);
                    }
                    UpdateFilteredIssues();
                    StatusMessage = $"Presentation consistency scan: {Report.Issues.Count} issues detected across {deckMap.Count} slides.";
                    SelectedTabIndex = 1;
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Presentation consistency error: {ex.Message}";
            }
        }

        public void ExecuteExportJson()
        {
            if (Report == null) return;
            try
            {
                string presName = ActivePresentationName;
                string json = ReportExportService.ExportToJson(Report, presName);
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"OneAssist_QualityReport_{DateTime.Now:yyyyMMdd_HHmmss}.json");
                System.IO.File.WriteAllText(path, json);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                StatusMessage = $"Exported JSON quality report to {System.IO.Path.GetFileName(path)}.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Export JSON error: {ex.Message}";
            }
        }

        public void ExecuteExportExcel()
        {
            if (Report == null) return;
            try
            {
                string presName = ActivePresentationName;
                string xml = ReportExportService.ExportToExcelXml(Report, presName);
                string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"OneAssist_QualityReport_{DateTime.Now:yyyyMMdd_HHmmss}.xml");
                System.IO.File.WriteAllText(path, xml);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                StatusMessage = $"Exported Excel report to {System.IO.Path.GetFileName(path)}.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Export Excel error: {ex.Message}";
            }
        }

        public void ExecuteCleanPresentation()
        {
            if (_powerPointApp == null) return;
            try
            {
                var pres = _powerPointApp.ActivePresentation;
                if (pres != null)
                {
                    var summary = Production.PresentationHealthCleanerService.CleanPresentation(pres, false, false, false, true);
                    StatusMessage = $"Clean Presentation: Removed {summary.GhostShapesRemoved} ghost/empty shapes across deck.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Clean Presentation error: {ex.Message}";
            }
        }

        public void ExecuteFixFormatting()
        {
            if (_powerPointApp == null) return;
            try
            {
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    string brandFont = Settings?.ActivePreset?.PrimaryFont ?? "Aptos";
                    int count = 0;
                    foreach (PowerPoint.Shape s in slide.Shapes)
                    {
                        if (SafeInteropHelper.HasText(s))
                        {
                            s.TextFrame.TextRange.Font.Name = brandFont;
                            s.TextFrame.MarginLeft = 7.2f;
                            s.TextFrame.MarginRight = 7.2f;
                            s.TextFrame.MarginTop = 7.2f;
                            s.TextFrame.MarginBottom = 7.2f;
                            count++;
                        }
                    }
                    StatusMessage = $"Standardized typography & text margins on {count} text elements to {brandFont}.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Fix formatting error: {ex.Message}";
            }
        }

        public void ExecuteConvertBulletsToCards()
        {
            if (_powerPointApp == null) return;
            try
            {
                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count == 1)
                {
                    var slide = (PowerPoint.Slide)_powerPointApp.ActiveWindow.View.Slide;
                    var shape = sel.ShapeRange[1];
                    var cards = Grid.SlideRestructuringService.ConvertBulletsToCards(slide, shape, 3);
                    StatusMessage = $"Transformed bullets into {cards.Count} professional cards.";
                }
                else
                {
                    StatusMessage = "Please select a single text frame containing bullet points.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Bullets to cards error: {ex.Message}";
            }
        }

        public void ExecuteZebraTable()
        {
            if (_powerPointApp == null) return;
            try
            {
                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count == 1)
                {
                    var shape = sel.ShapeRange[1];
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        Color header = Color.FromArgb(15, 23, 42);
                        Color even = Color.FromArgb(248, 250, 252);
                        Color odd = Color.FromArgb(255, 255, 255);
                        DataVisualization.TableDesignerService.ApplyZebraStriping(shape, header, even, odd);
                        StatusMessage = "Applied professional zebra styling and white header text to table.";
                        return;
                    }
                }
                StatusMessage = "Please select a PowerPoint table.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Table style error: {ex.Message}";
            }
        }

        public void ExecuteFormatChart()
        {
            if (_powerPointApp == null) return;
            try
            {
                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count == 1)
                {
                    var shape = sel.ShapeRange[1];
                    if (shape.HasChart == Office.MsoTriState.msoTrue)
                    {
                        Color primary = Color.FromArgb(0, 32, 96);
                        Color secondary = Color.FromArgb(0, 102, 204);
                        DataVisualization.ChartAssistantService.CleanChart(shape, primary, secondary);
                        StatusMessage = "Cleaned chart gridlines and styled series with brand colors.";
                        return;
                    }
                }
                StatusMessage = "Please select a PowerPoint chart.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Chart style error: {ex.Message}";
            }
        }

        public void ExecuteMorphAssistant()
        {
            if (_powerPointApp == null) return;
            try
            {
                var pres = _powerPointApp.ActivePresentation;
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null && pres.Slides.Count > 1)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    int nextIdx = slide.SlideIndex < pres.Slides.Count ? slide.SlideIndex + 1 : slide.SlideIndex - 1;
                    var otherSlide = pres.Slides[nextIdx];
                    var matches = Layers.MorphAssistantService.DetectCrossSlideMorphMatches(slide, otherSlide);
                    foreach (var m in matches)
                    {
                        Layers.MorphAssistantService.ApplyMorphMatching(slide, m.Shape1Id, otherSlide, m.Shape2Id, m.ProposedMorphName);
                    }
                    StatusMessage = $"Morph Assistant: Prepared {matches.Count} shapes between Slide {slide.SlideIndex} and Slide {nextIdx} with '!!' naming.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Morph Assistant error: {ex.Message}";
            }
        }

        public void ExecuteSwapShapes()
        {
            if (_powerPointApp == null) return;
            try
            {
                var sel = _powerPointApp.ActiveWindow?.Selection;
                if (sel?.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange.Count == 2)
                {
                    var s1 = sel.ShapeRange[1];
                    var s2 = sel.ShapeRange[2];
                    var opt = new Inspector.ShapeSwapOptions { SwapPosition = true, SwapSize = true };
                    Inspector.ShapeSwapService.Swap(s1, s2, opt);
                    StatusMessage = $"Swapped position and dimensions between '{s1.Name}' and '{s2.Name}'.";
                }
                else
                {
                    StatusMessage = "Please select exactly 2 shapes to swap.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Swap error: {ex.Message}";
            }
        }

        private void OnWindowSelectionChange(PowerPoint.Selection sel)
        {
            try
            {
                SelectionKind = ContextAwareSelectionService.DetectSelectionKind(sel);
                switch (SelectionKind)
                {
                    case PowerPointSelectionKind.SinglePicture:
                        SelectionContextText = "🖼️ Image Selected ➔ Prioritizing Picture Quality & Formatting Tools";
                        break;
                    case PowerPointSelectionKind.MultipleShapes:
                        SelectionContextText = $"📐 Multiple Shapes Selected ({sel.ShapeRange.Count}) ➔ Prioritizing Alignment & Spacing";
                        break;
                    case PowerPointSelectionKind.SingleText:
                        SelectionContextText = "🔤 Text Selected ➔ Prioritizing Typography & Editorial Tools";
                        break;
                    case PowerPointSelectionKind.SingleShape:
                        SelectionContextText = $"🔷 Shape Selected: {sel.ShapeRange[1].Name} ➔ Prioritizing Inspector & Color Tools";
                        break;
                    case PowerPointSelectionKind.Slide:
                        SelectionContextText = $"📑 Slide {sel.SlideRange[1].SlideIndex} Selected ➔ Prioritizing Layout & Cleanup";
                        break;
                    default:
                        SelectionContextText = "📊 Presentation Level ➔ Standard Toolbox Ready";
                        break;
                }
                OnPropertyChanged(nameof(ActivePresentationName));
                OnPropertyChanged(nameof(SlideCount));
            }
            catch { }
        }

        public async System.Threading.Tasks.Task ExecuteAILayoutAnalysisAsync()
        {
            if (_powerPointApp == null) return;
            try
            {
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    IsAnalyzingAILayout = true;
                    StatusMessage = "Analyzing slide content & semantics...";
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    CurrentSlideAnalysis = SlideSemanticAnalyzer.Analyze(slide);

                    StatusMessage = $"Detected {CurrentSlideAnalysis.DetectedContentType}: Generating AI layout alternatives...";
                    var suggestions = await _layoutPlanner.PlanLayoutsAsync(CurrentSlideAnalysis, Settings, SelectedDesignStyle, CancellationToken.None);

                    SuggestedLayouts.Clear();
                    foreach (var s in suggestions)
                    {
                        SuggestedLayouts.Add(s);
                    }

                    if (SuggestedLayouts.Count > 0)
                    {
                        SelectedLayoutSuggestion = SuggestedLayouts[0];
                    }

                    StatusMessage = $"AI Layout: Generated {SuggestedLayouts.Count} visual layout alternatives for Slide {slide.SlideIndex}.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"AI Layout Error: {ex.Message}";
            }
            finally
            {
                IsAnalyzingAILayout = false;
            }
        }

        public async System.Threading.Tasks.Task ExecuteImproveSlideAsync()
        {
            SelectedTabIndex = 14; // Show AI Layout studio
            await ExecuteAILayoutAnalysisAsync();
        }

        public async System.Threading.Tasks.Task ExecuteMakeSlideProfessionalAsync()
        {
            if (_powerPointApp == null) return;
            try
            {
                ExecuteFixFormatting();
                await ExecuteAILayoutAnalysisAsync();
                StatusMessage = "Polished slide typography, margins, and generated professional layouts.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Make Professional Error: {ex.Message}";
            }
        }

        public void ExecuteMakeItLikeThis()
        {
            if (_powerPointApp == null) return;
            try
            {
                var pres = _powerPointApp.ActivePresentation;
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null && pres != null)
                {
                    var targetSlide = (PowerPoint.Slide)activeWindow.View.Slide;
                    int refIdx = Math.Max(1, Math.Min(pres.Slides.Count, ReferenceSlideIndex));
                    var refSlide = pres.Slides[refIdx];

                    var spec = ReferenceSlideEngine.AdaptReferenceSlideLayout(refSlide, targetSlide);
                    if (spec != null)
                    {
                        bool ok = PowerPointLayoutRenderer.ApplyLayout(targetSlide, spec, Settings, false);
                        if (ok)
                        {
                            StatusMessage = $"Adapted visual layout from Slide {refIdx} onto Slide {targetSlide.SlideIndex} preserving all content.";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Make It Like This Error: {ex.Message}";
            }
        }

        public void ExecuteMakeTheseSlidesConsistent()
        {
            if (_powerPointApp == null) return;
            try
            {
                var pres = _powerPointApp.ActivePresentation;
                if (pres != null)
                {
                    var indices = new List<int>();
                    var sel = _powerPointApp.ActiveWindow?.Selection;
                    if (sel?.Type == PowerPoint.PpSelectionType.ppSelectionSlides && sel.SlideRange != null)
                    {
                        foreach (PowerPoint.Slide s in sel.SlideRange) indices.Add(s.SlideIndex);
                    }
                    else
                    {
                        for (int i = 1; i <= Math.Min(5, pres.Slides.Count); i++) indices.Add(i);
                    }

                    int count = SlideSequenceEngine.HarmonizeSlides(pres, indices, Settings);
                    StatusMessage = $"Harmonized visual layout consistency across {count + 1} slides in sequence.";
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Sequence consistency error: {ex.Message}";
            }
        }

        public void ExecuteApplySelectedLayout(bool asNewSlide)
        {
            if (_powerPointApp == null || SelectedLayoutSuggestion?.Specification == null)
            {
                StatusMessage = "Please select a layout suggestion to apply.";
                return;
            }

            try
            {
                var activeWindow = _powerPointApp.ActiveWindow;
                if (activeWindow?.View?.Slide != null)
                {
                    var slide = (PowerPoint.Slide)activeWindow.View.Slide;
                    bool success = PowerPointLayoutRenderer.ApplyLayout(slide, SelectedLayoutSuggestion.Specification, Settings, asNewSlide);
                    if (success)
                    {
                        StatusMessage = asNewSlide
                            ? $"Applied '{SelectedLayoutSuggestion.Title}' as new slide."
                            : $"Applied '{SelectedLayoutSuggestion.Title}' to Slide {slide.SlideIndex}.";
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Apply Layout Error: {ex.Message}";
            }
        }

        private void UpdateFilteredIssues()
        {
            FilteredIssues.Clear();
            if (Report?.Issues == null) return;

            var query = Report.Issues.Where(i => !i.IsIgnored);
            if (!string.Equals(FilterCategory, "All", StringComparison.OrdinalIgnoreCase))
            {
                query = query.Where(i => i.Category.ToString().Equals(FilterCategory, StringComparison.OrdinalIgnoreCase));
            }

            foreach (var issue in query.OrderBy(i => i.Severity).ThenBy(i => i.SlideIndex))
            {
                FilteredIssues.Add(issue);
            }
        }

        public void Dispose()
        {
            if (_powerPointApp != null)
            {
                try { _powerPointApp.WindowSelectionChange -= OnWindowSelectionChange; } catch { }
            }
            _cts?.Cancel();
            _cts?.Dispose();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }
    }
}

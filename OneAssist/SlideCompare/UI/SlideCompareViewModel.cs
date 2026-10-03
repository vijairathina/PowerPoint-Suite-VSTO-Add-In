using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using OneAssist.SlideCompare.Models;
using OneAssist.SlideCompare.Navigation;
using OneAssist.SlideCompare.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.SlideCompare.UI
{
    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Predicate<object> _canExecute;

        public RelayCommand(Action<object> execute, Predicate<object> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute == null || _canExecute(parameter);

        public void Execute(object parameter) => _execute(parameter);

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }

    public class SlideCompareViewModel : INotifyPropertyChanged
    {
        private readonly ComparisonService _comparisonService = new ComparisonService();
        private PowerPoint.Application _powerPointApp;
        private PowerPoint.Presentation _destinationPresentation;
        private CancellationTokenSource _cts;

        private string _sourceFilePath = string.Empty;
        private string _destinationPresentationName = "Active Presentation";
        private bool _isBusy;
        private int _progressPercentage;
        private string _progressText = "Ready";
        private bool _isSettingsOpen;
        private string _searchQuery = string.Empty;
        private System.Windows.Threading.DispatcherTimer _searchDebounceTimer;
        private bool _showOnlyIssues = false;
        private FilterMode _selectedFilter = FilterMode.All;

        private PresentationComparisonResult _comparisonResult;
        private ObservableCollection<SlideComparisonResult> _filteredResults = new ObservableCollection<SlideComparisonResult>();
        private CompareSettings _settings = new CompareSettings();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        public SlideCompareViewModel()
        {
            BrowseSourceCommand = new RelayCommand(_ => BrowseSource());
            StartCompareCommand = new RelayCommand(async _ => await RunCompareAsync(), _ => !IsBusy && !string.IsNullOrEmpty(SourceFilePath));
            CancelCompareCommand = new RelayCommand(_ => CancelCompare(), _ => IsBusy);
            SetFilterCommand = new RelayCommand(p => SetFilter(p?.ToString()));
            ToggleSettingsCommand = new RelayCommand(_ => IsSettingsOpen = !IsSettingsOpen);
            ToggleFilesSectionCommand = new RelayCommand(_ => IsFilesSectionVisible = !IsFilesSectionVisible);
            ToggleFilesVisibilityCommand = new RelayCommand(_ => IsFilesSectionVisible = !IsFilesSectionVisible);
            ToggleSummarySectionCommand = new RelayCommand(_ => IsSummarySectionCollapsed = !IsSummarySectionCollapsed);
            ToggleMaximizeResultsCommand = new RelayCommand(_ =>
            {
                IsMaximizedResults = !IsMaximizedResults;
                IsFilesSectionVisible = !IsMaximizedResults;
                IsSummarySectionCollapsed = IsMaximizedResults;
            });
            ExportHtmlCommand = new RelayCommand(_ => ExportReportHtml(), _ => ComparisonResult != null);
            ClearSearchCommand = new RelayCommand(_ => { SearchQuery = string.Empty; });
            NavigateToShapeCommand = new RelayCommand(p => NavigateToShape(p as ShapeComparisonResult));
            NavigateToSlideCommand = new RelayCommand(p => NavigateToSlide(p as SlideComparisonResult));

            _searchDebounceTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(200)
            };
            _searchDebounceTimer.Tick += (s, e) =>
            {
                _searchDebounceTimer.Stop();
                ApplyFilters();
            };
        }

        public void Initialize(PowerPoint.Application app)
        {
            _powerPointApp = app;
            RefreshDestination();
        }

        public void SetPresentations(PowerPoint.Presentation destPres, string sourcePath)
        {
            _destinationPresentation = destPres;
            if (_destinationPresentation != null)
            {
                try { DestinationPresentationName = _destinationPresentation.Name; } catch { }
            }
            SourceFilePath = sourcePath ?? string.Empty;
        }

        public void RefreshDestination()
        {
            try
            {
                if (_powerPointApp != null && _powerPointApp.Presentations.Count > 0)
                {
                    _destinationPresentation = _powerPointApp.ActivePresentation;
                    DestinationPresentationName = _destinationPresentation.Name;
                }
            }
            catch
            {
                DestinationPresentationName = "Active Presentation";
            }
        }

        #region Properties

        public string SourceFilePath
        {
            get => _sourceFilePath;
            set { _sourceFilePath = value; OnPropertyChanged(); OnPropertyChanged(nameof(SourceFileName)); }
        }

        public string SourceFileName => string.IsNullOrEmpty(_sourceFilePath) ? "No file selected" : Path.GetFileName(_sourceFilePath);

        public string DestinationPresentationName
        {
            get => _destinationPresentationName;
            set { _destinationPresentationName = value; OnPropertyChanged(); }
        }

        public bool IsBusy
        {
            get => _isBusy;
            set { _isBusy = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNotBusy)); }
        }

        public bool IsNotBusy => !IsBusy;

        public int ProgressPercentage
        {
            get => _progressPercentage;
            set { _progressPercentage = value; OnPropertyChanged(); }
        }

        public string ProgressText
        {
            get => _progressText;
            set { _progressText = value; OnPropertyChanged(); }
        }

        public bool IsSettingsOpen
        {
            get => _isSettingsOpen;
            set { _isSettingsOpen = value; OnPropertyChanged(); }
        }

        public string SearchQuery
        {
            get => _searchQuery;
            set
            {
                if (_searchQuery != value)
                {
                    _searchQuery = value ?? string.Empty;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(HasSearchQuery));

                    if (string.IsNullOrWhiteSpace(_searchQuery))
                    {
                        _searchDebounceTimer?.Stop();
                        ApplyFilters();
                    }
                    else
                    {
                        _searchDebounceTimer?.Stop();
                        _searchDebounceTimer?.Start();
                    }
                }
            }
        }

        public bool HasSearchQuery => !string.IsNullOrEmpty(_searchQuery);

        public bool ShowOnlyIssues
        {
            get => _showOnlyIssues;
            set
            {
                _showOnlyIssues = value;
                OnPropertyChanged();
                if (ComparisonResult?.SlideResults != null)
                {
                    foreach (var s in ComparisonResult.SlideResults)
                    {
                        s.ShowOnlyIssues = value;
                    }
                }
                if (value && _selectedFilter == FilterMode.All)
                    SelectedFilter = FilterMode.IssuesOnly;
                else if (!value && _selectedFilter == FilterMode.IssuesOnly)
                    SelectedFilter = FilterMode.All;
                ApplyFilters();
            }
        }

        public FilterMode SelectedFilter
        {
            get => _selectedFilter;
            set { _selectedFilter = value; OnPropertyChanged(); ApplyFilters(); }
        }

        public CompareSettings Settings
        {
            get => _settings;
            set { _settings = value; OnPropertyChanged(); }
        }

        public PresentationComparisonResult ComparisonResult
        {
            get => _comparisonResult;
            private set
            {
                _comparisonResult = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SummaryMatched));
                OnPropertyChanged(nameof(SummaryModified));
                OnPropertyChanged(nameof(SummaryAdded));
                OnPropertyChanged(nameof(SummaryDeleted));
                OnPropertyChanged(nameof(SummaryMoved));
                OnPropertyChanged(nameof(SummaryReordered));
                OnPropertyChanged(nameof(SummarySplit));
                OnPropertyChanged(nameof(SummaryMerged));
                OnPropertyChanged(nameof(TotalIssuesCount));
                OnPropertyChanged(nameof(HasResults));
                ApplyFilters();
            }
        }

        public ObservableCollection<SlideComparisonResult> FilteredResults
        {
            get => _filteredResults;
            set { _filteredResults = value; OnPropertyChanged(); }
        }

        public bool HasResults => ComparisonResult != null;

        private bool _isFilesSectionVisible = true;
        public bool IsFilesSectionVisible
        {
            get => _isFilesSectionVisible;
            set
            {
                _isFilesSectionVisible = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(FilesButtonTooltip));
            }
        }
        public string FilesButtonTooltip => IsFilesSectionVisible ? "Hide Presentations & Settings Panel" : "Show Presentations & Settings Panel";

        private bool _isSummarySectionCollapsed;
        public bool IsSummarySectionCollapsed
        {
            get => _isSummarySectionCollapsed;
            set
            {
                _isSummarySectionCollapsed = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsSummarySectionExpanded));
            }
        }
        public bool IsSummarySectionExpanded => !IsSummarySectionCollapsed;

        private bool _isMaximizedResults;
        public bool IsMaximizedResults
        {
            get => _isMaximizedResults;
            set
            {
                _isMaximizedResults = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(MaximizeButtonIcon));
                OnPropertyChanged(nameof(MaximizeButtonText));
                OnPropertyChanged(nameof(MaximizeButtonTooltip));
            }
        }

        public string MaximizeButtonIcon => IsMaximizedResults ? "🗗" : "🗖";
        public string MaximizeButtonText => IsMaximizedResults ? "🗗" : "🗖";
        public string MaximizeButtonTooltip => IsMaximizedResults ? "Restore View (Show files & summary panels)" : "Maximize Results (Hide files & summary panels)";

        public int SummaryMatched => ComparisonResult != null
            ? (ComparisonResult.MatchedContentCount > 0 ? ComparisonResult.MatchedContentCount : ComparisonResult.MatchedCount)
            : 0;

        public int SummaryMatchedContent => ComparisonResult?.MatchedContentCount ?? 0;
        public int SummaryMatchedSlides => ComparisonResult?.MatchedSlidesCount ?? (ComparisonResult?.MatchedCount ?? 0);
        public string SummaryMatchedTooltip => $"{SummaryMatched} matched items ({SummaryMatchedContent} content shapes, {SummaryMatchedSlides} slides)";

        public int SummaryModified => ComparisonResult?.ModifiedCount ?? 0;
        public int SummaryAdded => ComparisonResult?.AddedCount ?? 0;
        public int SummaryDeleted => ComparisonResult?.DeletedCount ?? 0;
        public int SummaryMoved => ComparisonResult?.MovedCount ?? 0;
        public int SummaryReordered => ComparisonResult?.ReorderedCount ?? 0;
        public int SummarySplit => ComparisonResult?.SplitCount ?? 0;
        public int SummaryMerged => ComparisonResult?.MergedCount ?? 0;
        public int TotalIssuesCount => ComparisonResult?.TotalIssuesCount ?? 0;

        #endregion

        #region Commands

        public ICommand BrowseSourceCommand { get; }
        public ICommand StartCompareCommand { get; }
        public ICommand CancelCompareCommand { get; }
        public ICommand SetFilterCommand { get; }
        public ICommand ToggleSettingsCommand { get; }
        public ICommand ToggleFilesSectionCommand { get; }
        public ICommand ToggleFilesVisibilityCommand { get; }
        public ICommand ToggleSummarySectionCommand { get; }
        public ICommand ToggleMaximizeResultsCommand { get; }
        public ICommand ExportHtmlCommand { get; }
        public ICommand ClearSearchCommand { get; }
        public ICommand NavigateToShapeCommand { get; }
        public ICommand NavigateToSlideCommand { get; }

        #endregion

        #region Methods

        private void BrowseSource()
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Source PowerPoint Presentation";
                ofd.Filter = "PowerPoint Presentations (*.pptx;*.pptm;*.ppt)|*.pptx;*.pptm;*.ppt|All Files (*.*)|*.*";
                ofd.CheckFileExists = true;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    SourceFilePath = ofd.FileName;
                }
            }
        }

        public async Task RunCompareAsync()
        {
            if (string.IsNullOrEmpty(SourceFilePath) || !File.Exists(SourceFilePath))
            {
                System.Windows.MessageBox.Show("Please select a valid source PowerPoint presentation file.", "Slide Compare", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            RefreshDestination();
            if (_destinationPresentation == null)
            {
                System.Windows.MessageBox.Show("No active destination PowerPoint presentation found. Please open a presentation first.", "Slide Compare", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            IsBusy = true;
            ProgressPercentage = 0;
            ProgressText = "Starting comparison...";
            _cts = new CancellationTokenSource();

            var progress = new Progress<Tuple<int, string>>(p =>
            {
                ProgressPercentage = p.Item1;
                ProgressText = p.Item2;
            });

            try
            {
                var result = await _comparisonService.RunComparisonAsync(
                    _powerPointApp,
                    _destinationPresentation,
                    SourceFilePath,
                    Settings,
                    progress,
                    _cts.Token);

                ComparisonResult = result;
                IsFilesSectionVisible = false; // Once compare runs, hide it!
                ProgressPercentage = 100;
                ProgressText = $"Comparison complete: {result.TotalIssuesCount} issues found.";
            }
            catch (OperationCanceledException)
            {
                ProgressText = "Comparison was cancelled.";
            }
            catch (Exception ex)
            {
                ProgressText = "Error during comparison.";
                System.Windows.MessageBox.Show($"Failed to complete comparison: {ex.Message}", "Slide Compare Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                _cts?.Dispose();
                _cts = null;
            }
        }

        private void CancelCompare()
        {
            _cts?.Cancel();
            ProgressText = "Cancelling...";
        }

        public void SetFilter(string filterName)
        {
            if (Enum.TryParse<FilterMode>(filterName, true, out var mode))
            {
                SelectedFilter = mode;
                ShowOnlyIssues = (mode != FilterMode.All && mode != FilterMode.Matched);
            }
        }

        public void ApplyFilters()
        {
            if (ComparisonResult?.SlideResults == null)
            {
                FilteredResults = new ObservableCollection<SlideComparisonResult>();
                return;
            }

            // When Matched or All is selected, ensure matched content is visible in slide cards
            bool forceShowAll = (SelectedFilter == FilterMode.Matched || SelectedFilter == FilterMode.All);
            foreach (var s in ComparisonResult.SlideResults)
            {
                s.ShowOnlyIssues = forceShowAll ? false : ShowOnlyIssues;
            }

            IEnumerable<SlideComparisonResult> query = ComparisonResult.SlideResults;

            // Apply Type Filter
            switch (SelectedFilter)
            {
                case FilterMode.IssuesOnly:
                    query = query.Where(s => s.HasIssues);
                    break;
                case FilterMode.Matched:
                    query = query.Where(s => s.MatchType == SlideMatchType.Matched 
                        || (s.ShapeResults != null && s.ShapeResults.Any(sh => sh.IsMatched)));
                    break;
                case FilterMode.Modified:
                    query = query.Where(s => s.MatchType == SlideMatchType.Modified || (s.ShapeResults != null && s.ShapeResults.Any(sh => sh.IsIssue)));
                    break;
                case FilterMode.Added:
                    query = query.Where(s => s.MatchType == SlideMatchType.Added 
                        || (s.ShapeResults != null && s.ShapeResults.Any(sh => sh.MatchType == ShapeMatchType.Added || sh.IsAdded || (sh.DiffTokens != null && sh.DiffTokens.Any(t => t.DiffType == TextDiffType.Added)))));
                    break;
                case FilterMode.Deleted:
                    query = query.Where(s => s.MatchType == SlideMatchType.Deleted 
                        || (s.ShapeResults != null && s.ShapeResults.Any(sh => sh.MatchType == ShapeMatchType.Deleted || sh.IsDeleted || (sh.DiffTokens != null && sh.DiffTokens.Any(t => t.DiffType == TextDiffType.Deleted)))));
                    break;
                case FilterMode.Moved:
                    query = query.Where(s => s.MatchType == SlideMatchType.Moved || (s.ShapeResults != null && s.ShapeResults.Any(sh => sh.MatchType == ShapeMatchType.Moved)));
                    break;
                case FilterMode.Reordered:
                    query = query.Where(s => s.MatchType == SlideMatchType.Reordered);
                    break;
                case FilterMode.Split:
                    query = query.Where(s => s.MatchType == SlideMatchType.Split);
                    break;
                case FilterMode.Merged:
                    query = query.Where(s => s.MatchType == SlideMatchType.Merged);
                    break;
                case FilterMode.TextChanges:
                    query = query.Where(s => s.ShapeResults != null && s.ShapeResults.Any(sh => sh.HasTextChanges));
                    break;
                case FilterMode.ShapeChanges:
                    query = query.Where(s => s.ShapeResults != null && s.ShapeResults.Any(sh => sh.MatchType == ShapeMatchType.Added || sh.MatchType == ShapeMatchType.Deleted));
                    break;
                case FilterMode.FormattingChanges:
                    query = query.Where(s => s.ShapeResults != null && s.ShapeResults.Any(sh => sh.HasFormattingChanges));
                    break;
                case FilterMode.All:
                default:
                    break;
            }

            // Apply Text Search Query (Optimized cached lookup)
            if (!string.IsNullOrWhiteSpace(SearchQuery))
            {
                string q = SearchQuery.Trim().ToLowerInvariant();
                query = query.Where(s => s.MatchesSearch(q));
            }

            FilteredResults = new ObservableCollection<SlideComparisonResult>(query);
        }

        private void NavigateToShape(ShapeComparisonResult shape)
        {
            if (shape == null || _destinationPresentation == null) return;

            int targetSlide = shape.DestinationSlideNumber > 0 ? shape.DestinationSlideNumber : shape.SourceSlideNumber;
            PowerPointNavigationService.NavigateToShape(
                _destinationPresentation,
                targetSlide,
                shape.DestinationShapeName ?? shape.SourceShapeName,
                shape.DestinationShapeId,
                shape.DestinationShapeIndex);
        }

        private void NavigateToSlide(SlideComparisonResult slide)
        {
            if (slide == null || _destinationPresentation == null) return;

            int targetSlide = slide.PrimaryDestinationSlide > 0 ? slide.PrimaryDestinationSlide : slide.PrimarySourceSlide;
            PowerPointNavigationService.NavigateToSlide(_destinationPresentation, targetSlide);
        }

        private void ExportReportHtml()
        {
            if (ComparisonResult == null) return;

            using (var sfd = new SaveFileDialog())
            {
                sfd.Title = "Save Slide Compare HTML Report";
                sfd.Filter = "HTML Files (*.html)|*.html";
                sfd.FileName = $"SlideCompare_{DateTime.Now:yyyyMMdd_HHmmss}.html";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string html = _comparisonService.ExportToHtml(ComparisonResult);
                    File.WriteAllText(sfd.FileName, html);
                    System.Windows.MessageBox.Show("HTML report exported successfully.", "Slide Compare", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        #endregion
    }
}

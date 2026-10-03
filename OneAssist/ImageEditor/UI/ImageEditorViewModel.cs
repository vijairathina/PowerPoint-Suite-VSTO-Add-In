using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using OneAssist.ImageEditor.Core;
using OneAssist.ImageEditor.PowerPointServices;
using OneAssist.ImageEditor.Processing;
using OneAssist.ImageEditor.Settings;
using OneAssist.ImageEditor.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.ImageEditor.UI
{
    public class ImageEditorViewModel : INotifyPropertyChanged, IDisposable
    {
        private readonly ImageDocument _document;
        private readonly IBackgroundRemovalService _bgRemovalService;
        private readonly IObjectRemovalService _objectRemovalService;
        private readonly ShapeSelectionService _selectionService;
        private PowerPoint.Application _powerPointApp;

        private BitmapSource _currentPreviewSource;
        private BitmapSource _originalPreviewSource;
        private ToolType _activeTool = ToolType.RemoveBg;
        private bool _isProcessing;
        private int _processingProgress;
        private string _processingStatusText;
        private bool _isPictureSelected;
        private string _selectionStatusMessage = "Please select an image in PowerPoint.";
        private bool _showBeforeAfterComparison;
        private double _comparisonSplitPercent = 50.0;

        // Brush settings
        private int _brushSize = 35;
        private int _brushHardness = 80;
        private int _brushOpacity = 100;
        private int _brushFeather = 4;
        private bool _isRestoreMode;

        // Background removal options
        private int _bgTolerance = 38;
        private int _bgFeather = 3;

        // Background configuration
        private BackgroundType _selectedBackgroundType = BackgroundType.Transparent;
        private System.Windows.Media.Color _customBackgroundColor = System.Windows.Media.Colors.White;
        private float _gradientAngle = 45f;
        private System.Windows.Media.Color _gradientColor1 = System.Windows.Media.Color.FromRgb(99, 102, 241);
        private System.Windows.Media.Color _gradientColor2 = System.Windows.Media.Color.FromRgb(168, 85, 247);

        // Crop configuration
        private CropAspectRatio _cropRatio = CropAspectRatio.Free;
        private Rect _cropRectangle;

        // FX settings
        private FxSettings _fxSettings = new FxSettings();

        // Zoom & Pan
        private double _zoom = 1.0;

        public event PropertyChangedEventHandler PropertyChanged;

        public ImageDocument Document => _document;

        public ImageEditorViewModel()
        {
            _document = new ImageDocument();
            _document.DocumentChanged += (s, e) => UpdatePreview();
            _document.History.HistoryChanged += (s, e) =>
            {
                OnPropertyChanged(nameof(CanUndo));
                OnPropertyChanged(nameof(CanRedo));
            };

            _bgRemovalService = new SmartBackgroundRemovalService();
            _objectRemovalService = new FastMarchingInpaintingService();
            _selectionService = new ShapeSelectionService();
            _selectionService.SelectionChanged += OnSelectionChanged;

            InitializeCommands();

            try
            {
                var app = GetPowerPointApp();
                if (app != null)
                {
                    _selectionService.Initialize(app);
                }
            }
            catch { }
        }

        public PowerPoint.Application GetPowerPointApp()
        {
            if (_powerPointApp != null) return _powerPointApp;
            try
            {
                _powerPointApp = Globals.ThisAddIn?.Application;
            }
            catch { }
            return _powerPointApp;
        }

        public void Initialize(PowerPoint.Application app)
        {
            _powerPointApp = app ?? GetPowerPointApp();
            if (_powerPointApp != null)
            {
                _selectionService.Initialize(_powerPointApp);
                CheckSelection();
            }
        }

        #region Properties

        public BitmapSource CurrentPreviewSource
        {
            get => _currentPreviewSource;
            private set { _currentPreviewSource = value; OnPropertyChanged(); }
        }

        public BitmapSource OriginalPreviewSource
        {
            get => _originalPreviewSource;
            private set { _originalPreviewSource = value; OnPropertyChanged(); }
        }

        public bool HasImage => _document.HasImage;
        public int ImageWidth => _document.Width;
        public int ImageHeight => _document.Height;

        public bool IsPictureSelected
        {
            get => _isPictureSelected;
            set { _isPictureSelected = value; OnPropertyChanged(); }
        }

        public string SelectionStatusMessage
        {
            get => _selectionStatusMessage;
            set { _selectionStatusMessage = value; OnPropertyChanged(); }
        }

        public ToolType ActiveTool
        {
            get => _activeTool;
            set
            {
                if (_activeTool != value)
                {
                    _activeTool = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsRemoveBgToolActive));
                    OnPropertyChanged(nameof(IsMarkRemoveToolActive));
                    OnPropertyChanged(nameof(IsEraserToolActive));
                    OnPropertyChanged(nameof(IsBackgroundToolActive));
                    OnPropertyChanged(nameof(IsCropToolActive));
                    OnPropertyChanged(nameof(IsAdjustToolActive));
                    OnPropertyChanged(nameof(IsTransformToolActive));
                    OnPropertyChanged(nameof(IsFxToolActive));
                }
            }
        }

        public bool IsRemoveBgToolActive  => ActiveTool == ToolType.RemoveBg;
        public bool IsMarkRemoveToolActive => ActiveTool == ToolType.MarkRemove;
        public bool IsEraserToolActive     => ActiveTool == ToolType.Eraser || ActiveTool == ToolType.Restore;
        public bool IsBackgroundToolActive => ActiveTool == ToolType.Background;
        public bool IsCropToolActive       => ActiveTool == ToolType.Crop;
        public bool IsAdjustToolActive     => ActiveTool == ToolType.Adjust;
        public bool IsTransformToolActive  => ActiveTool == ToolType.Transform;
        public bool IsFxToolActive         => ActiveTool == ToolType.FX;

        public bool IsProcessing
        {
            get => _isProcessing;
            set { _isProcessing = value; OnPropertyChanged(); }
        }

        public int ProcessingProgress
        {
            get => _processingProgress;
            set { _processingProgress = value; OnPropertyChanged(); }
        }

        public string ProcessingStatusText
        {
            get => _processingStatusText;
            set { _processingStatusText = value; OnPropertyChanged(); }
        }

        public bool ShowBeforeAfterComparison
        {
            get => _showBeforeAfterComparison;
            set { _showBeforeAfterComparison = value; OnPropertyChanged(); }
        }

        public double ComparisonSplitPercent
        {
            get => _comparisonSplitPercent;
            set { _comparisonSplitPercent = value; OnPropertyChanged(); }
        }

        public bool CanUndo => _document.History.CanUndo;
        public bool CanRedo => _document.History.CanRedo;

        public int BrushSize
        {
            get => _brushSize;
            set { _brushSize = value; OnPropertyChanged(); }
        }

        public int BrushHardness
        {
            get => _brushHardness;
            set { _brushHardness = value; OnPropertyChanged(); }
        }

        public int BrushOpacity
        {
            get => _brushOpacity;
            set { _brushOpacity = value; OnPropertyChanged(); }
        }

        public int BrushFeather
        {
            get => _brushFeather;
            set { _brushFeather = value; OnPropertyChanged(); }
        }

        public bool IsRestoreMode
        {
            get => _isRestoreMode;
            set
            {
                _isRestoreMode = value;
                ActiveTool = value ? ToolType.Restore : ToolType.Eraser;
                OnPropertyChanged();
            }
        }

        public int BgTolerance
        {
            get => _bgTolerance;
            set { _bgTolerance = value; OnPropertyChanged(); }
        }

        public int BgFeather
        {
            get => _bgFeather;
            set { _bgFeather = value; OnPropertyChanged(); }
        }

        public BackgroundType SelectedBackgroundType
        {
            get => _selectedBackgroundType;
            set
            {
                _selectedBackgroundType = value;
                _document.Background.Type = value;
                OnPropertyChanged();
                UpdatePreview();
            }
        }

        public System.Windows.Media.Color CustomBackgroundColor
        {
            get => _customBackgroundColor;
            set
            {
                _customBackgroundColor = value;
                _document.Background.CustomColor = System.Drawing.Color.FromArgb(value.A, value.R, value.G, value.B);
                OnPropertyChanged();
                UpdatePreview();
            }
        }

        public float GradientAngle
        {
            get => _gradientAngle;
            set
            {
                _gradientAngle = value;
                _document.Background.GradientAngle = value;
                OnPropertyChanged();
                UpdatePreview();
            }
        }

        // Adjustments proxy properties
        public int Brightness
        {
            get => _document.Adjustments.Brightness;
            set { _document.Adjustments.Brightness = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int Contrast
        {
            get => _document.Adjustments.Contrast;
            set { _document.Adjustments.Contrast = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int Saturation
        {
            get => _document.Adjustments.Saturation;
            set { _document.Adjustments.Saturation = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int Exposure
        {
            get => _document.Adjustments.Exposure;
            set { _document.Adjustments.Exposure = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int Sharpness
        {
            get => _document.Adjustments.Sharpness;
            set { _document.Adjustments.Sharpness = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int Blur
        {
            get => _document.Adjustments.Blur;
            set { _document.Adjustments.Blur = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int ImageOpacity
        {
            get => _document.Adjustments.Opacity;
            set { _document.Adjustments.Opacity = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public CropAspectRatio CropRatio
        {
            get => _cropRatio;
            set { _cropRatio = value; OnPropertyChanged(); }
        }

        public Rect CropRectangle
        {
            get => _cropRectangle;
            set { _cropRectangle = value; OnPropertyChanged(); }
        }

        // FX properties
        public FxSettings Fx => _fxSettings;

        public bool FxDropShadowEnabled
        {
            get => _fxSettings.DropShadowEnabled;
            set { _fxSettings.DropShadowEnabled = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxShadowOffsetX
        {
            get => _fxSettings.ShadowOffsetX;
            set { _fxSettings.ShadowOffsetX = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxShadowOffsetY
        {
            get => _fxSettings.ShadowOffsetY;
            set { _fxSettings.ShadowOffsetY = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxShadowBlur
        {
            get => _fxSettings.ShadowBlur;
            set { _fxSettings.ShadowBlur = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxShadowOpacityPct
        {
            get => (int)(_fxSettings.ShadowOpacity * 100);
            set { _fxSettings.ShadowOpacity = value / 100f; OnPropertyChanged(); UpdatePreview(); }
        }

        public bool FxGlowEnabled
        {
            get => _fxSettings.GlowEnabled;
            set { _fxSettings.GlowEnabled = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxGlowRadius
        {
            get => _fxSettings.GlowRadius;
            set { _fxSettings.GlowRadius = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxGlowIntensityPct
        {
            get => (int)(_fxSettings.GlowIntensity * 100);
            set { _fxSettings.GlowIntensity = value / 100f; OnPropertyChanged(); UpdatePreview(); }
        }

        public bool FxVignetteEnabled
        {
            get => _fxSettings.VignetteEnabled;
            set { _fxSettings.VignetteEnabled = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxVignetteStrengthPct
        {
            get => (int)(_fxSettings.VignetteStrength * 100);
            set { _fxSettings.VignetteStrength = value / 100f; OnPropertyChanged(); UpdatePreview(); }
        }

        public bool FxTintEnabled
        {
            get => _fxSettings.TintEnabled;
            set { _fxSettings.TintEnabled = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxTintStrengthPct
        {
            get => (int)(_fxSettings.TintStrength * 100);
            set { _fxSettings.TintStrength = value / 100f; OnPropertyChanged(); UpdatePreview(); }
        }

        public int FxColorTemperature
        {
            get => _fxSettings.ColorTemperature;
            set { _fxSettings.ColorTemperature = Math.Max(-100, Math.Min(100, value)); OnPropertyChanged(); UpdatePreview(); }
        }

        public double Zoom
        {
            get => _zoom;
            set { _zoom = Math.Max(0.1, Math.Min(5.0, value)); OnPropertyChanged(); }
        }

        #endregion

        #region Commands

        public ICommand UndoCommand { get; private set; }
        public ICommand RedoCommand { get; private set; }
        public ICommand ResetCommand { get; private set; }
        public ICommand CancelCommand { get; private set; }
        public ICommand ApplyToPowerPointCommand { get; private set; }
        public ICommand LoadSelectedImageCommand { get; private set; }
        public ICommand RemoveBackgroundCommand { get; private set; }
        public ICommand RemoveMarkedAreaCommand { get; private set; }
        public ICommand ClearMarkingsCommand { get; private set; }
        public ICommand RotateLeftCommand { get; private set; }
        public ICommand RotateRightCommand { get; private set; }
        public ICommand FlipHorizontalCommand { get; private set; }
        public ICommand FlipVerticalCommand { get; private set; }
        public ICommand ApplyCropCommand { get; private set; }
        public ICommand CancelCropCommand { get; private set; }
        public ICommand ResetAdjustmentsCommand { get; private set; }
        public ICommand SelectBackgroundImageCommand { get; private set; }
        public ICommand ToggleBeforeAfterCommand { get; private set; }
        public ICommand FitToPanelCommand { get; private set; }
        public ICommand Zoom100Command { get; private set; }

        private void InitializeCommands()
        {
            UndoCommand = new RelayCommand(() => _document.Undo(), () => CanUndo && !IsProcessing);
            RedoCommand = new RelayCommand(() => _document.Redo(), () => CanRedo && !IsProcessing);
            ResetCommand = new RelayCommand(() => _document.ResetToOriginal(), () => HasImage && !IsProcessing);
            CancelCommand = new RelayCommand(CancelEdits, () => !IsProcessing);
            ApplyToPowerPointCommand = new RelayCommand(async () => await ApplyToPowerPointAsync(), () => HasImage && !IsProcessing);
            LoadSelectedImageCommand = new RelayCommand(LoadSelectedPicture, () => !IsProcessing);

            RemoveBackgroundCommand = new RelayCommand(async () => await ExecuteRemoveBackgroundAsync(), () => HasImage && !IsProcessing);
            RemoveMarkedAreaCommand = new RelayCommand(async () => await ExecuteInpaintMarkedAsync(), () => HasImage && !IsProcessing);
            ClearMarkingsCommand = new RelayCommand(ClearMaskMarkings, () => HasImage && !IsProcessing);

            RotateLeftCommand = new RelayCommand(() => RotateImage(RotateFlipType.Rotate270FlipNone), () => HasImage && !IsProcessing);
            RotateRightCommand = new RelayCommand(() => RotateImage(RotateFlipType.Rotate90FlipNone), () => HasImage && !IsProcessing);
            FlipHorizontalCommand = new RelayCommand(() => RotateImage(RotateFlipType.RotateNoneFlipX), () => HasImage && !IsProcessing);
            FlipVerticalCommand = new RelayCommand(() => RotateImage(RotateFlipType.RotateNoneFlipY), () => HasImage && !IsProcessing);

            ApplyCropCommand = new RelayCommand(ExecuteApplyCrop, () => HasImage && !IsProcessing);
            CancelCropCommand = new RelayCommand(() => ActiveTool = ToolType.Select, () => !IsProcessing);
            ResetAdjustmentsCommand = new RelayCommand(() => { _document.Adjustments.Reset(); RefreshAdjustmentProperties(); UpdatePreview(); }, () => HasImage && !IsProcessing);

            SelectBackgroundImageCommand = new RelayCommand(ChooseCustomBackgroundImage, () => HasImage && !IsProcessing);
            ToggleBeforeAfterCommand = new RelayCommand(() => ShowBeforeAfterComparison = !ShowBeforeAfterComparison, () => HasImage);
            FitToPanelCommand = new RelayCommand(() => Zoom = 1.0);
            Zoom100Command = new RelayCommand(() => Zoom = 1.0);
        }

        #endregion

        #region Operations

        public void CheckSelection()
        {
            var app = GetPowerPointApp();
            bool isPic = PowerPointImageService.IsPictureSelected(app, out string reason);
            IsPictureSelected = isPic;
            SelectionStatusMessage = reason;

            if (isPic && !HasImage)
            {
                LoadSelectedPicture();
            }
        }

        private void OnSelectionChanged(object sender, ShapeSelectionEventArgs e)
        {
            IsPictureSelected = e.IsPictureSelected;
            SelectionStatusMessage = e.Message;

            // If user selected an image and we don't have one open, auto-load it
            if (e.IsPictureSelected && !HasImage)
            {
                try
                {
                    var dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher;
                    if (dispatcher != null && !dispatcher.CheckAccess())
                    {
                        dispatcher.Invoke(new Action(LoadSelectedPicture));
                    }
                    else
                    {
                        LoadSelectedPicture();
                    }
                }
                catch
                {
                    LoadSelectedPicture();
                }
            }
        }

        public void LoadSelectedPicture()
        {
            var app = GetPowerPointApp();
            if (app == null)
            {
                SelectionStatusMessage = "PowerPoint application is unavailable.";
                return;
            }

            var res = PowerPointImageService.ExtractSelectedImage(app);
            if (res.Success && res.Image != null)
            {
                _document.LoadNewImage(res.Image, res.Metadata);
                OriginalPreviewSource = BitmapHelper.ToBitmapSource(res.Image);
                IsPictureSelected = true;
                SelectionStatusMessage = string.Empty;
                ActiveTool = ToolType.RemoveBg;
                Zoom = 1.0;
                UpdatePreview();
            }
            else
            {
                IsPictureSelected = false;
                SelectionStatusMessage = res.ErrorMessage ?? "Please select an image in PowerPoint.";
            }
        }

        private async Task ExecuteRemoveBackgroundAsync()
        {
            if (!HasImage || IsProcessing) return;

            IsProcessing = true;
            ProcessingStatusText = "Removing background...";
            ProcessingProgress = 10;

            try
            {
                var progress = new Progress<int>(p => ProcessingProgress = p);
                var options = new BackgroundRemovalOptions
                {
                    Tolerance = BgTolerance,
                    Feather = BgFeather
                };

                var cleanImage = await _bgRemovalService.RemoveBackgroundAsync(_document.WorkingImage, options, progress);
                if (cleanImage != null)
                {
                    _document.UpdateWorkingImage(cleanImage, "Remove Background");
                    _document.UpdateMask(ImageMask.FromBitmapAlpha(cleanImage));
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Background removal error: {ex.Message}", "Image Editor", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                IsProcessing = false;
                ProcessingStatusText = string.Empty;
            }
        }

        private async Task ExecuteInpaintMarkedAsync()
        {
            if (!HasImage || IsProcessing) return;
            if (_document.RemovalMask == null) return;

            // Check if any pixels are marked
            bool hasMarkedPixels = false;
            for (int i = 0; i < _document.RemovalMask.Data.Length; i++)
            {
                if (_document.RemovalMask.Data[i] > 20)
                {
                    hasMarkedPixels = true;
                    break;
                }
            }

            if (!hasMarkedPixels)
            {
                System.Windows.MessageBox.Show("Please paint over an unwanted object with the brush first.", "Mark & Remove", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            IsProcessing = true;
            ProcessingStatusText = "Removing marked object...";
            ProcessingProgress = 10;

            try
            {
                var progress = new Progress<int>(p => ProcessingProgress = p);
                var options = new InpaintOptions { Radius = 10 };

                var cleaned = await _objectRemovalService.RemoveObjectAsync(_document.WorkingImage, _document.RemovalMask, options, progress);
                if (cleaned != null)
                {
                    _document.UpdateWorkingImage(cleaned, "Mark & Remove Object");
                    _document.ClearRemovalMask();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Object removal error: {ex.Message}", "Image Editor", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                IsProcessing = false;
                ProcessingStatusText = string.Empty;
            }
        }

        private void ClearMaskMarkings()
        {
            _document.ClearRemovalMask();
        }

        private void RotateImage(RotateFlipType type)
        {
            if (!HasImage) return;

            var rotated = BitmapHelper.TransformBitmap(_document.WorkingImage, type);
            if (rotated != null)
            {
                _document.UpdateWorkingImage(rotated, "Rotate/Flip");
            }
        }

        private void ExecuteApplyCrop()
        {
            if (!HasImage || CropRectangle.Width <= 0 || CropRectangle.Height <= 0) return;

            var cropRect = new Rectangle(
                (int)Math.Round(CropRectangle.X),
                (int)Math.Round(CropRectangle.Y),
                (int)Math.Round(CropRectangle.Width),
                (int)Math.Round(CropRectangle.Height));

            var cropped = BitmapHelper.CropBitmap(_document.WorkingImage, cropRect);
            if (cropped != null)
            {
                if (_document.OriginalImage != null)
                {
                    var origCropped = BitmapHelper.CropBitmap(_document.OriginalImage, cropRect);
                    OriginalPreviewSource = BitmapHelper.ToBitmapSource(origCropped);
                    origCropped.Dispose();
                }

                _document.UpdateWorkingImage(cropped, "Crop");
                ActiveTool = ToolType.Select;
            }
        }

        private void ChooseCustomBackgroundImage()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select Background Image";
                dialog.Filter = "Image Files (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|All Files (*.*)|*.*";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var loaded = BitmapHelper.LoadBitmapFromDiskSafely(dialog.FileName);
                    if (loaded != null)
                    {
                        _document.Background.BackgroundImage?.Dispose();
                        _document.Background.BackgroundImage = loaded;
                        SelectedBackgroundType = BackgroundType.Image;
                        UpdatePreview();
                    }
                }
            }
        }

        private void CancelEdits()
        {
            _document.ResetToOriginal();
            ActiveTool = ToolType.Select;
        }

        private async Task ApplyToPowerPointAsync()
        {
            if (!HasImage || IsProcessing) return;

            IsProcessing = true;
            ProcessingStatusText = "Applying to PowerPoint...";
            ProcessingProgress = 30;

            try
            {
                // Render final composite for export
                var finalBitmap = await Task.Run(() =>
                {
                    var bmp = BackgroundRenderer.RenderComposite(
                        _document.WorkingImage,
                        _document.AlphaMask,
                        _document.Adjustments,
                        _document.Background,
                        drawCheckerboardForTransparency: false);

                    // Apply FX effects in sequence
                    if (_fxSettings.ColorTemperature != 0)
                    {
                        var tmp = DesignerFxService.ApplyColorTemperature(bmp, _fxSettings.ColorTemperature);
                        bmp.Dispose(); bmp = tmp;
                    }
                    if (_fxSettings.TintEnabled)
                    {
                        var tmp = DesignerFxService.ApplyColorTint(bmp, _fxSettings.TintColor, _fxSettings.TintStrength);
                        bmp.Dispose(); bmp = tmp;
                    }
                    if (_fxSettings.VignetteEnabled)
                    {
                        var tmp = DesignerFxService.ApplyVignette(bmp, _fxSettings.VignetteStrength);
                        bmp.Dispose(); bmp = tmp;
                    }
                    if (_fxSettings.GlowEnabled)
                    {
                        var tmp = DesignerFxService.ApplyGlow(bmp, _fxSettings.GlowRadius, _fxSettings.GlowIntensity, _fxSettings.GlowColor);
                        bmp.Dispose(); bmp = tmp;
                    }
                    if (_fxSettings.DropShadowEnabled)
                    {
                        var tmp = DesignerFxService.ApplyDropShadow(bmp, _fxSettings.ShadowOffsetX, _fxSettings.ShadowOffsetY, _fxSettings.ShadowBlur, _fxSettings.ShadowOpacity, _fxSettings.ShadowColor);
                        bmp.Dispose(); bmp = tmp;
                    }

                    return bmp;
                });

                ProcessingProgress = 70;

                bool success = PowerPointApplyService.ApplyImageToPowerPoint(
                    GetPowerPointApp(),
                    finalBitmap,
                    _document.Metadata,
                    out string error);

                finalBitmap?.Dispose();

                if (success)
                {
                    ProcessingProgress = 100;
                    System.Windows.MessageBox.Show("Image successfully applied to PowerPoint slide!", "OneAssist Image Editor", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    System.Windows.MessageBox.Show($"Failed to apply image: {error}", "OneAssist Image Editor", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error applying to slide: {ex.Message}", "OneAssist Image Editor", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsProcessing = false;
                ProcessingStatusText = string.Empty;
            }
        }

        public void UpdatePreview()
        {
            if (!HasImage)
            {
                CurrentPreviewSource = null;
                return;
            }

            try
            {
                // Composite with checkerboard for transparency preview in UI
                using (var previewBitmap = BackgroundRenderer.RenderComposite(
                    _document.WorkingImage,
                    _document.AlphaMask,
                    _document.Adjustments,
                    _document.Background,
                    drawCheckerboardForTransparency: true))
                {
                    CurrentPreviewSource = BitmapHelper.ToBitmapSource(previewBitmap);
                }

                OnPropertyChanged(nameof(HasImage));
                OnPropertyChanged(nameof(ImageWidth));
                OnPropertyChanged(nameof(ImageHeight));
                OnPropertyChanged(nameof(CanUndo));
                OnPropertyChanged(nameof(CanRedo));
            }
            catch
            {
                // Suppress rendering preview issues
            }
        }

        private void RefreshAdjustmentProperties()
        {
            OnPropertyChanged(nameof(Brightness));
            OnPropertyChanged(nameof(Contrast));
            OnPropertyChanged(nameof(Saturation));
            OnPropertyChanged(nameof(Exposure));
            OnPropertyChanged(nameof(Sharpness));
            OnPropertyChanged(nameof(Blur));
            OnPropertyChanged(nameof(ImageOpacity));
        }

        #endregion

        protected void OnPropertyChanged([CallerMemberName] string propName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }

        public void Dispose()
        {
            _document.Dispose();
            _selectionService.Dispose();
        }
    }
}

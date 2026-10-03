using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using OneAssist.ImageEditor.Core;
using OneAssist.ImageEditor.Settings;

namespace OneAssist.ImageEditor.UI
{
    public partial class ImageEditorControl : System.Windows.Controls.UserControl
    {
        public ImageEditorViewModel ViewModel => DataContext as ImageEditorViewModel;

        public ImageEditorControl()
        {
            InitializeComponent();
            KeyDown += OnImageEditorKeyDown;
            DataContextChanged += (s, e) =>
            {
                if (CanvasCtrl != null && e.NewValue is ImageEditorViewModel vm)
                {
                    CanvasCtrl.ViewModel = vm;
                }
            };

            // Re-wire shortcuts whenever settings change
            SettingsManager.Instance.SettingsChanged += (s, e) => { /* shortcuts handled via KeyDown */ };
        }

        private void OnImageEditorKeyDown(object sender, KeyEventArgs e)
        {
            if (ViewModel == null) return;

            var cfg = SettingsManager.Instance.Current;

            // Helper: check if a shortcut string matches the current key event
            bool Matches(string shortcut)
            {
                if (!SettingsManager.ParseShortcut(shortcut, out ModifierKeys mods, out Key k)) return false;
                return e.Key == k && Keyboard.Modifiers == mods;
            }

            if (Matches(cfg.ShortcutUndo) || (e.Key == Key.Z && Keyboard.Modifiers == ModifierKeys.Control))
            {
                if (ViewModel.UndoCommand.CanExecute(null)) ViewModel.UndoCommand.Execute(null);
                e.Handled = true; return;
            }
            if (Matches(cfg.ShortcutRedo) || (e.Key == Key.Y && Keyboard.Modifiers == ModifierKeys.Control)
                || (e.Key == Key.Z && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift)))
            {
                if (ViewModel.RedoCommand.CanExecute(null)) ViewModel.RedoCommand.Execute(null);
                e.Handled = true; return;
            }
            if (Matches(cfg.ShortcutApply))
            {
                if (ViewModel.ApplyToPowerPointCommand.CanExecute(null)) ViewModel.ApplyToPowerPointCommand.Execute(null);
                e.Handled = true; return;
            }
            if (Matches(cfg.ShortcutRemoveBg))
            {
                if (ViewModel.RemoveBackgroundCommand.CanExecute(null)) ViewModel.RemoveBackgroundCommand.Execute(null);
                e.Handled = true; return;
            }
            if (Matches(cfg.ShortcutReload))
            {
                if (ViewModel.LoadSelectedImageCommand.CanExecute(null)) ViewModel.LoadSelectedImageCommand.Execute(null);
                e.Handled = true; return;
            }
            if (Matches(cfg.ShortcutToggleCompare))
            {
                if (ViewModel.ToggleBeforeAfterCommand.CanExecute(null)) ViewModel.ToggleBeforeAfterCommand.Execute(null);
                e.Handled = true; return;
            }

            // Tool switcher keys
            if (Matches(cfg.ShortcutToolRemoveBg))  { ViewModel.ActiveTool = ToolType.RemoveBg;  e.Handled = true; return; }
            if (Matches(cfg.ShortcutToolMarkRemove)) { ViewModel.ActiveTool = ToolType.MarkRemove; e.Handled = true; return; }
            if (Matches(cfg.ShortcutToolEraser))     { ViewModel.ActiveTool = ToolType.Eraser;    e.Handled = true; return; }
            if (Matches(cfg.ShortcutToolBackground)) { ViewModel.ActiveTool = ToolType.Background; e.Handled = true; return; }
            if (Matches(cfg.ShortcutToolCrop))       { ViewModel.ActiveTool = ToolType.Crop;       e.Handled = true; return; }
            if (Matches(cfg.ShortcutToolAdjust))     { ViewModel.ActiveTool = ToolType.Adjust;     e.Handled = true; return; }
            if (Matches(cfg.ShortcutToolTransform))  { ViewModel.ActiveTool = ToolType.Transform;  e.Handled = true; return; }
        }

        private void OnSelectRemoveBgTool(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ActiveTool = ToolType.RemoveBg;
        }

        private void OnSelectMarkRemoveTool(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ActiveTool = ToolType.MarkRemove;
        }

        private void OnSelectEraserTool(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ActiveTool = ViewModel.IsRestoreMode ? ToolType.Restore : ToolType.Eraser;
        }

        private void OnSelectBackgroundTool(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ActiveTool = ToolType.Background;
        }

        private void OnSelectCropTool(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null)
            {
                ViewModel.ActiveTool = ToolType.Crop;
                if (ViewModel.Document != null && ViewModel.Document.HasImage)
                {
                    ViewModel.CropRectangle = new Rect(0, 0, ViewModel.Document.Width, ViewModel.Document.Height);
                }
            }
        }

        private void OnSelectAdjustTool(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ActiveTool = ToolType.Adjust;
        }

        private void OnSelectTransformTool(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ActiveTool = ToolType.Transform;
        }

        private void OnEraseModeChecked(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.IsRestoreMode = false;
        }

        private void OnRestoreModeChecked(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.IsRestoreMode = true;
        }

        private void OnBgTransparentClick(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.SelectedBackgroundType = BackgroundType.Transparent;
        }

        private void OnBgWhiteClick(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.SelectedBackgroundType = BackgroundType.White;
        }

        private void OnBgBlackClick(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.SelectedBackgroundType = BackgroundType.Black;
        }

        private void OnBgLightGrayClick(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.SelectedBackgroundType = BackgroundType.LightGray;
        }

        private void OnBgGradientClick(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.SelectedBackgroundType = BackgroundType.LinearGradient;
        }

        private void OnCropFreeClick(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.CropRatio = CropAspectRatio.Free;
        }

        private void OnCrop1x1Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.CropRatio = CropAspectRatio.Ratio1x1;
        }

        private void OnCrop4x3Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.CropRatio = CropAspectRatio.Ratio4x3;
        }

        private void OnCrop16x9Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.CropRatio = CropAspectRatio.Ratio16x9;
        }

        private void OnSelectFxTool(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ActiveTool = ToolType.FX;
        }

        private void OnOpenSettingsClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var window = new SettingsWindow();
                window.Owner = System.Windows.Window.GetWindow(this);
                bool? result = window.ShowDialog();

                if (result == true && ViewModel != null)
                {
                    // Apply new defaults to the ViewModel
                    var cfg = SettingsManager.Instance.Current;
                    ViewModel.BgTolerance = cfg.DefaultBgTolerance;
                    ViewModel.BgFeather   = cfg.DefaultBgFeather;
                    ViewModel.BrushSize   = cfg.DefaultBrushSize;
                    ViewModel.BrushHardness = cfg.DefaultBrushHardness;
                    ViewModel.BrushOpacity  = cfg.DefaultBrushOpacity;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Could not open settings: {ex.Message}",
                    "Settings",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
        }
    }
}

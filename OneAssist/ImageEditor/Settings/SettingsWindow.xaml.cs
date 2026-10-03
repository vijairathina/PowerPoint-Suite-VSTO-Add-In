using System;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace OneAssist.ImageEditor.Settings
{
    /// <summary>
    /// Code-behind for the Settings dialog.
    /// </summary>
    public partial class SettingsWindow : Window
    {
        private EditorSettings _editCopy;

        public SettingsWindow()
        {
            InitializeComponent();
            _editCopy = CloneSettings(SettingsManager.Instance.Current);
            PopulateFields(_editCopy);
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Populate UI fields from a settings object
        // ─────────────────────────────────────────────────────────────────────────────
        private void PopulateFields(EditorSettings s)
        {
            TxtUndo.Text            = s.ShortcutUndo;
            TxtRedo.Text            = s.ShortcutRedo;
            TxtApply.Text           = s.ShortcutApply;
            TxtRemoveBg.Text        = s.ShortcutRemoveBg;
            TxtReload.Text          = s.ShortcutReload;
            TxtToggleCompare.Text   = s.ShortcutToggleCompare;
            TxtToolRemoveBg.Text    = s.ShortcutToolRemoveBg;
            TxtToolMarkRemove.Text  = s.ShortcutToolMarkRemove;
            TxtToolEraser.Text      = s.ShortcutToolEraser;
            TxtToolBackground.Text  = s.ShortcutToolBackground;
            TxtToolCrop.Text        = s.ShortcutToolCrop;
            TxtToolAdjust.Text      = s.ShortcutToolAdjust;
            TxtToolTransform.Text   = s.ShortcutToolTransform;

            TxtBgTolerance.Text     = s.DefaultBgTolerance.ToString();
            TxtBgFeather.Text       = s.DefaultBgFeather.ToString();

            TxtBrushSize.Text       = s.DefaultBrushSize.ToString();
            TxtBrushHardness.Text   = s.DefaultBrushHardness.ToString();
            TxtBrushOpacity.Text    = s.DefaultBrushOpacity.ToString();

            ChkAutoLoad.IsChecked   = s.AutoLoadOnSelection;
            ChkAutoDetect.IsChecked = s.AutoDetectOnShow;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Keyboard shortcut capture: intercept key down, build shortcut string
        // ─────────────────────────────────────────────────────────────────────────────
        private void OnShortcutKeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
            var tb = sender as System.Windows.Controls.TextBox;
            if (tb == null) return;

            // Ignore modifier-only key presses
            if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl ||
                e.Key == Key.LeftAlt  || e.Key == Key.RightAlt  ||
                e.Key == Key.LeftShift|| e.Key == Key.RightShift||
                e.Key == Key.LWin    || e.Key == Key.RWin)
            {
                return;
            }

            var sb = new StringBuilder();
            if ((Keyboard.Modifiers & ModifierKeys.Control) != 0) sb.Append("Ctrl+");
            if ((Keyboard.Modifiers & ModifierKeys.Alt)     != 0) sb.Append("Alt+");
            if ((Keyboard.Modifiers & ModifierKeys.Shift)   != 0) sb.Append("Shift+");

            // Get actual key name
            Key actualKey = e.Key == Key.System ? e.SystemKey : e.Key;
            sb.Append(actualKey.ToString());

            tb.Text = sb.ToString();
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Collect UI fields back into a settings object
        // ─────────────────────────────────────────────────────────────────────────────
        private EditorSettings CollectFields()
        {
            var s = new EditorSettings();

            s.ShortcutUndo            = TxtUndo.Text.Trim();
            s.ShortcutRedo            = TxtRedo.Text.Trim();
            s.ShortcutApply           = TxtApply.Text.Trim();
            s.ShortcutRemoveBg        = TxtRemoveBg.Text.Trim();
            s.ShortcutReload          = TxtReload.Text.Trim();
            s.ShortcutToggleCompare   = TxtToggleCompare.Text.Trim();
            s.ShortcutToolRemoveBg    = TxtToolRemoveBg.Text.Trim();
            s.ShortcutToolMarkRemove  = TxtToolMarkRemove.Text.Trim();
            s.ShortcutToolEraser      = TxtToolEraser.Text.Trim();
            s.ShortcutToolBackground  = TxtToolBackground.Text.Trim();
            s.ShortcutToolCrop        = TxtToolCrop.Text.Trim();
            s.ShortcutToolAdjust      = TxtToolAdjust.Text.Trim();
            s.ShortcutToolTransform   = TxtToolTransform.Text.Trim();

            s.DefaultBgTolerance      = ParseInt(TxtBgTolerance.Text, 38, 10, 90);
            s.DefaultBgFeather        = ParseInt(TxtBgFeather.Text, 3, 0, 20);
            s.DefaultBrushSize        = ParseInt(TxtBrushSize.Text, 35, 1, 200);
            s.DefaultBrushHardness    = ParseInt(TxtBrushHardness.Text, 80, 0, 100);
            s.DefaultBrushOpacity     = ParseInt(TxtBrushOpacity.Text, 100, 1, 100);

            s.AutoLoadOnSelection     = ChkAutoLoad.IsChecked == true;
            s.AutoDetectOnShow        = ChkAutoDetect.IsChecked == true;

            return s;
        }

        private static int ParseInt(string text, int defaultVal, int min, int max)
        {
            if (int.TryParse(text, out int v))
                return Math.Max(min, Math.Min(max, v));
            return defaultVal;
        }

        private static EditorSettings CloneSettings(EditorSettings src)
        {
            // Simple clone via JSON round-trip using the DataContract serializer
            var s = new EditorSettings();
            s.ShortcutUndo            = src.ShortcutUndo;
            s.ShortcutRedo            = src.ShortcutRedo;
            s.ShortcutApply           = src.ShortcutApply;
            s.ShortcutRemoveBg        = src.ShortcutRemoveBg;
            s.ShortcutReload          = src.ShortcutReload;
            s.ShortcutToggleCompare   = src.ShortcutToggleCompare;
            s.ShortcutToolRemoveBg    = src.ShortcutToolRemoveBg;
            s.ShortcutToolMarkRemove  = src.ShortcutToolMarkRemove;
            s.ShortcutToolEraser      = src.ShortcutToolEraser;
            s.ShortcutToolBackground  = src.ShortcutToolBackground;
            s.ShortcutToolCrop        = src.ShortcutToolCrop;
            s.ShortcutToolAdjust      = src.ShortcutToolAdjust;
            s.ShortcutToolTransform   = src.ShortcutToolTransform;
            s.DefaultBgTolerance      = src.DefaultBgTolerance;
            s.DefaultBgFeather        = src.DefaultBgFeather;
            s.DefaultBrushSize        = src.DefaultBrushSize;
            s.DefaultBrushHardness    = src.DefaultBrushHardness;
            s.DefaultBrushOpacity     = src.DefaultBrushOpacity;
            s.AutoLoadOnSelection     = src.AutoLoadOnSelection;
            s.AutoDetectOnShow        = src.AutoDetectOnShow;
            s.TaskPaneWidth           = src.TaskPaneWidth;
            return s;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // Button handlers
        // ─────────────────────────────────────────────────────────────────────────────
        private void OnSaveClick(object sender, RoutedEventArgs e)
        {
            var newSettings = CollectFields();
            SettingsManager.Instance.Apply(newSettings);
            DialogResult = true;
            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void OnResetClick(object sender, RoutedEventArgs e)
        {
            var confirm = MessageBox.Show(
                "Reset all settings to their defaults?",
                "Reset Settings",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm == MessageBoxResult.Yes)
            {
                PopulateFields(EditorSettings.Default);
            }
        }
    }
}

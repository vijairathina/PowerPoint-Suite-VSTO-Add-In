using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Windows.Input;

namespace OneAssist.ImageEditor.Settings
{
    /// <summary>
    /// Persisted user settings for the Image Editor. Stored as JSON in AppData\Local\OneAssist\.
    /// </summary>
    [DataContract]
    public class EditorSettings
    {
        // ── Keyboard Shortcuts ──────────────────────────────────────────────────────
        [DataMember] public string ShortcutUndo { get; set; }            = "Ctrl+Z";
        [DataMember] public string ShortcutRedo { get; set; }            = "Ctrl+Y";
        [DataMember] public string ShortcutApply { get; set; }           = "Ctrl+S";
        [DataMember] public string ShortcutRemoveBg { get; set; }        = "Ctrl+B";
        [DataMember] public string ShortcutReload { get; set; }          = "Ctrl+R";
        [DataMember] public string ShortcutToggleCompare { get; set; }   = "Ctrl+D";
        [DataMember] public string ShortcutToolRemoveBg { get; set; }    = "1";
        [DataMember] public string ShortcutToolMarkRemove { get; set; }  = "2";
        [DataMember] public string ShortcutToolEraser { get; set; }      = "3";
        [DataMember] public string ShortcutToolBackground { get; set; }  = "4";
        [DataMember] public string ShortcutToolCrop { get; set; }        = "5";
        [DataMember] public string ShortcutToolAdjust { get; set; }      = "6";
        [DataMember] public string ShortcutToolTransform { get; set; }   = "7";

        // ── Background removal defaults ─────────────────────────────────────────────
        [DataMember] public int DefaultBgTolerance { get; set; }         = 38;
        [DataMember] public int DefaultBgFeather { get; set; }           = 3;

        // ── Brush defaults ─────────────────────────────────────────────────────────
        [DataMember] public int DefaultBrushSize { get; set; }           = 35;
        [DataMember] public int DefaultBrushHardness { get; set; }       = 80;
        [DataMember] public int DefaultBrushOpacity { get; set; }        = 100;

        // ── Canvas preferences ─────────────────────────────────────────────────────
        [DataMember] public bool AutoLoadOnSelection { get; set; }       = true;
        [DataMember] public bool AutoDetectOnShow { get; set; }          = true;
        [DataMember] public int TaskPaneWidth { get; set; }              = 420;

        // ── Export ─────────────────────────────────────────────────────────────────
        [DataMember] public bool ExportAsPng { get; set; }               = true;
        [DataMember] public int ExportJpegQuality { get; set; }          = 95;

        public static EditorSettings Default => new EditorSettings();
    }

    /// <summary>
    /// Singleton settings manager that loads/saves settings to AppData\Local\OneAssist\settings.json.
    /// </summary>
    public class SettingsManager
    {
        private static SettingsManager _instance;
        public static SettingsManager Instance => _instance ?? (_instance = new SettingsManager());

        private static readonly string SettingsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OneAssist");
        private static readonly string SettingsPath = Path.Combine(SettingsDir, "settings.json");

        private EditorSettings _current;
        public EditorSettings Current => _current ?? (_current = Load());

        public event EventHandler SettingsChanged;

        private EditorSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var bytes = File.ReadAllBytes(SettingsPath);
                    using (var ms = new MemoryStream(bytes))
                    {
                        var ser = new DataContractJsonSerializer(typeof(EditorSettings));
                        return (EditorSettings)ser.ReadObject(ms) ?? EditorSettings.Default;
                    }
                }
            }
            catch { }
            return EditorSettings.Default;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);
                using (var ms = new MemoryStream())
                {
                    var ser = new DataContractJsonSerializer(typeof(EditorSettings));
                    ser.WriteObject(ms, Current);
                    File.WriteAllBytes(SettingsPath, ms.ToArray());
                }
            }
            catch { }
        }

        public void Apply(EditorSettings newSettings)
        {
            _current = newSettings;
            Save();
            SettingsChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Reset()
        {
            _current = EditorSettings.Default;
            Save();
            SettingsChanged?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Parses a shortcut string like "Ctrl+Z" or "Alt+Shift+B" into ModifierKeys and Key.
        /// Returns false if the string is empty or unparseable.
        /// </summary>
        public static bool ParseShortcut(string shortcut, out ModifierKeys mods, out Key key)
        {
            mods = ModifierKeys.None;
            key = Key.None;
            if (string.IsNullOrWhiteSpace(shortcut)) return false;

            var parts = shortcut.Split('+');
            string keyPart = parts[parts.Length - 1].Trim();

            foreach (var part in parts)
            {
                var p = part.Trim().ToLower();
                if (p == "ctrl" || p == "control") mods |= ModifierKeys.Control;
                else if (p == "alt") mods |= ModifierKeys.Alt;
                else if (p == "shift") mods |= ModifierKeys.Shift;
                else if (p == "win" || p == "windows") mods |= ModifierKeys.Windows;
            }

            if (Enum.TryParse(keyPart, true, out Key parsedKey) && parsedKey != Key.None)
            {
                key = parsedKey;
                return true;
            }

            return false;
        }
    }
}

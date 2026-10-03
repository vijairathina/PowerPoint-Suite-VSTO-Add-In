using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace OneAssist.Assistant.Settings
{
    [DataContract]
    public class BrandColorDefinition
    {
        [DataMember] public string Name { get; set; }
        [DataMember] public string HexColor { get; set; }
        [DataMember] public string Role { get; set; } // Primary, Secondary, Accent, Background, Text

        public Color ToColor()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(HexColor)) return Color.Empty;
                return ColorTranslator.FromHtml(HexColor);
            }
            catch
            {
                return Color.Empty;
            }
        }
    }

    [DataContract]
    public class BrandPreset
    {
        [DataMember] public string PresetName { get; set; } = "Corporate Standard";
        [DataMember] public string PrimaryFont { get; set; } = "Aptos";
        [DataMember] public string HeadingFont { get; set; } = "Aptos Display";
        [DataMember] public float MinTitleFontSize { get; set; } = 24f;
        [DataMember] public float MinBodyFontSize { get; set; } = 12f;
        [DataMember] public List<BrandColorDefinition> Colors { get; set; } = new List<BrandColorDefinition>();

        public static BrandPreset CreateDefault()
        {
            return new BrandPreset
            {
                PresetName = "Default Corporate",
                PrimaryFont = "Aptos",
                HeadingFont = "Aptos Display",
                MinTitleFontSize = 24f,
                MinBodyFontSize = 12f,
                Colors = new List<BrandColorDefinition>
                {
                    new BrandColorDefinition { Name = "Brand Navy", HexColor = "#002060", Role = "Primary" },
                    new BrandColorDefinition { Name = "Accent Teal", HexColor = "#00A88F", Role = "Accent" },
                    new BrandColorDefinition { Name = "Dark Charcoal", HexColor = "#262626", Role = "Text" },
                    new BrandColorDefinition { Name = "Clean White", HexColor = "#FFFFFF", Role = "Background" },
                    new BrandColorDefinition { Name = "Light Gray", HexColor = "#F2F4F7", Role = "Secondary" }
                }
            };
        }
    }

    [DataContract]
    public class AssistantSettings
    {
        // ── Rule Activation ──────────────────────────────────────────
        [DataMember] public bool EnableFontConsistencyCheck { get; set; } = true;
        [DataMember] public bool EnableFontSizeCheck { get; set; } = true;
        [DataMember] public bool EnableMarginCheck { get; set; } = true;
        [DataMember] public bool EnableAlignmentCheck { get; set; } = true;
        [DataMember] public bool EnableSpacingCheck { get; set; } = true;
        [DataMember] public bool EnableColorContrastCheck { get; set; } = true;
        [DataMember] public bool EnableBrandColorCheck { get; set; } = true;
        [DataMember] public bool EnableObjectOverlapCheck { get; set; } = true;
        [DataMember] public bool EnableOffSlideCheck { get; set; } = true;

        // ── Tolerances & Thresholds ──────────────────────────────────
        [DataMember] public float AlmostAlignedTolerancePoints { get; set; } = 3.0f; // Within 3pt is considered "almost aligned"
        [DataMember] public float SpacingInconsistencyTolerancePoints { get; set; } = 2.0f;
        [DataMember] public double MinWcagContrastRatio { get; set; } = 4.5; // WCAG AA
        [DataMember] public int MaxColorsPerSlideThreshold { get; set; } = 6;
        [DataMember] public float SnapGridSizePoints { get; set; } = 10.0f;

        // ── Brand Presets ────────────────────────────────────────────
        [DataMember] public BrandPreset ActivePreset { get; set; } = BrandPreset.CreateDefault();
        [DataMember] public List<BrandPreset> SavedPresets { get; set; } = new List<BrandPreset>();

        // ── Editorial Terminology (Corporate Dictionary) ─────────────
        [DataMember] public Dictionary<string, string> TerminologyReplacements { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Power Point", "PowerPoint" },
            { "PowerBI", "Power BI" },
            { "API's", "APIs" },
            { "e-mail", "email" },
            { "datas", "data" }
        };

        public static AssistantSettings Default => new AssistantSettings();
    }

    public class AssistantSettingsManager
    {
        private static AssistantSettingsManager _instance;
        public static AssistantSettingsManager Instance => _instance ?? (_instance = new AssistantSettingsManager());

        private static readonly string SettingsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OneAssist");
        private static readonly string SettingsFilePath =
            Path.Combine(SettingsDir, "assistant_settings.json");

        private AssistantSettings _current;
        public AssistantSettings Current => _current ?? (_current = Load());

        public event EventHandler SettingsChanged;

        private AssistantSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    var bytes = File.ReadAllBytes(SettingsFilePath);
                    using (var ms = new MemoryStream(bytes))
                    {
                        var ser = new DataContractJsonSerializer(typeof(AssistantSettings));
                        var settings = (AssistantSettings)ser.ReadObject(ms);
                        if (settings != null) return settings;
                    }
                }
            }
            catch { }

            var def = AssistantSettings.Default;
            return def;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);
                using (var ms = new MemoryStream())
                {
                    var ser = new DataContractJsonSerializer(typeof(AssistantSettings));
                    ser.WriteObject(ms, Current);
                    File.WriteAllBytes(SettingsFilePath, ms.ToArray());
                }
            }
            catch { }
        }

        public void Apply(AssistantSettings newSettings)
        {
            _current = newSettings;
            Save();
            SettingsChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}

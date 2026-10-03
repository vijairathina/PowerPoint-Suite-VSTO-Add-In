using System;
using System.IO;
using System.Runtime.Serialization.Json;
using OneAssist.Assistant.Settings;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.Services
{
    public static class PresetManager
    {
        public static string ExportPresetToJson(BrandPreset preset)
        {
            if (preset == null) return string.Empty;
            using (var ms = new MemoryStream())
            {
                var ser = new DataContractJsonSerializer(typeof(BrandPreset));
                ser.WriteObject(ms, preset);
                return System.Text.Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        public static BrandPreset ImportPresetFromJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(json);
            using (var ms = new MemoryStream(bytes))
            {
                var ser = new DataContractJsonSerializer(typeof(BrandPreset));
                return (BrandPreset)ser.ReadObject(ms);
            }
        }

        public static void ApplyBrandPresetToSlide(PowerPoint.Slide slide, BrandPreset preset)
        {
            if (slide == null || preset == null) return;
            try
            {
                foreach (PowerPoint.Shape s in slide.Shapes)
                {
                    if (SafeInteropHelper.HasText(s))
                    {
                        var tf = s.TextFrame.TextRange;
                        float? sz = SafeInteropHelper.GetFontSize(s);

                        // If title
                        if (sz.HasValue && sz.Value >= preset.MinTitleFontSize)
                        {
                            if (!string.IsNullOrEmpty(preset.HeadingFont))
                            {
                                tf.Font.Name = preset.HeadingFont;
                            }
                        }
                        else
                        {
                            // Body text
                            if (!string.IsNullOrEmpty(preset.PrimaryFont))
                            {
                                tf.Font.Name = preset.PrimaryFont;
                            }
                        }
                    }
                }
            }
            catch { }
        }
    }
}

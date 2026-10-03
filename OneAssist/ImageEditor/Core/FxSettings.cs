using System.Drawing;

namespace OneAssist.ImageEditor.Core
{
    /// <summary>
    /// Stores designer FX effect parameters: drop shadow, glow, vignette, color tint, and color temperature.
    /// </summary>
    public class FxSettings
    {
        // Drop Shadow
        public bool DropShadowEnabled { get; set; }  = false;
        public int ShadowOffsetX     { get; set; }   = 8;
        public int ShadowOffsetY     { get; set; }   = 8;
        public int ShadowBlur        { get; set; }   = 12;
        public float ShadowOpacity   { get; set; }   = 0.55f;
        public Color ShadowColor     { get; set; }   = Color.Black;

        // Glow
        public bool GlowEnabled      { get; set; }   = false;
        public int GlowRadius        { get; set; }   = 10;
        public float GlowIntensity   { get; set; }   = 0.5f;
        public Color GlowColor       { get; set; }   = Color.White;

        // Vignette
        public bool VignetteEnabled  { get; set; }   = false;
        public float VignetteStrength { get; set; }  = 0.5f;

        // Color Tint
        public bool TintEnabled      { get; set; }   = false;
        public Color TintColor       { get; set; }   = Color.FromArgb(100, 150, 200);
        public float TintStrength    { get; set; }   = 0.3f;

        // Color Temperature (-100 = cool, +100 = warm)
        public int ColorTemperature  { get; set; }   = 0;

        public void Reset()
        {
            DropShadowEnabled = false;
            ShadowOffsetX     = 8;
            ShadowOffsetY     = 8;
            ShadowBlur        = 12;
            ShadowOpacity     = 0.55f;
            ShadowColor       = Color.Black;

            GlowEnabled       = false;
            GlowRadius        = 10;
            GlowIntensity     = 0.5f;
            GlowColor         = Color.White;

            VignetteEnabled   = false;
            VignetteStrength  = 0.5f;

            TintEnabled       = false;
            TintColor         = Color.FromArgb(100, 150, 200);
            TintStrength      = 0.3f;

            ColorTemperature  = 0;
        }
    }
}

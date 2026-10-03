using System;

namespace OneAssist.ImageEditor.Core
{
    public class AdjustmentSettings
    {
        public int Brightness { get; set; } = 0;   // -100 to 100
        public int Contrast { get; set; } = 0;     // -100 to 100
        public int Saturation { get; set; } = 0;   // -100 to 100
        public int Exposure { get; set; } = 0;     // -100 to 100
        public int Sharpness { get; set; } = 0;    // 0 to 100
        public int Blur { get; set; } = 0;         // 0 to 20
        public int Opacity { get; set; } = 100;    // 0 to 100

        public bool IsDefault =>
            Brightness == 0 &&
            Contrast == 0 &&
            Saturation == 0 &&
            Exposure == 0 &&
            Sharpness == 0 &&
            Blur == 0 &&
            Opacity == 100;

        public void Reset()
        {
            Brightness = 0;
            Contrast = 0;
            Saturation = 0;
            Exposure = 0;
            Sharpness = 0;
            Blur = 0;
            Opacity = 100;
        }

        public AdjustmentSettings Clone()
        {
            return new AdjustmentSettings
            {
                Brightness = this.Brightness,
                Contrast = this.Contrast,
                Saturation = this.Saturation,
                Exposure = this.Exposure,
                Sharpness = this.Sharpness,
                Blur = this.Blur,
                Opacity = this.Opacity
            };
        }
    }
}

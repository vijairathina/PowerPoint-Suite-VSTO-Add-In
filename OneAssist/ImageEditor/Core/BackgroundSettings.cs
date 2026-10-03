using System.Drawing;
using OneAssist.ImageEditor.Utils;

namespace OneAssist.ImageEditor.Core
{
    public class BackgroundSettings
    {
        public BackgroundType Type { get; set; } = BackgroundType.Transparent;
        public Color CustomColor { get; set; } = Color.White;
        public GradientType GradientType { get; set; } = GradientType.Linear;
        public float GradientAngle { get; set; } = 45f;
        public Color GradientColor1 { get; set; } = Color.FromArgb(99, 102, 241); // Indigo
        public Color GradientColor2 { get; set; } = Color.FromArgb(168, 85, 247); // Purple
        public Bitmap BackgroundImage { get; set; }
        public float ImageScale { get; set; } = 1.0f;
        public float ImageOffsetX { get; set; } = 0f;
        public float ImageOffsetY { get; set; } = 0f;

        public BackgroundSettings Clone()
        {
            return new BackgroundSettings
            {
                Type = this.Type,
                CustomColor = this.CustomColor,
                GradientType = this.GradientType,
                GradientAngle = this.GradientAngle,
                GradientColor1 = this.GradientColor1,
                GradientColor2 = this.GradientColor2,
                BackgroundImage = this.BackgroundImage != null ? BitmapHelper.Clone32bppArgb(this.BackgroundImage) : null,
                ImageScale = this.ImageScale,
                ImageOffsetX = this.ImageOffsetX,
                ImageOffsetY = this.ImageOffsetY
            };
        }
    }
}

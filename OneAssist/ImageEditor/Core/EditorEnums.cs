namespace OneAssist.ImageEditor.Core
{
    public enum ToolType
    {
        None,
        Select,
        RemoveBg,
        MarkRemove,
        Eraser,
        Restore,
        Background,
        Crop,
        Adjust,
        Transform,
        FX
    }

    public enum BackgroundType
    {
        Transparent,
        White,
        Black,
        LightGray,
        CustomColor,
        LinearGradient,
        RadialGradient,
        Image
    }

    public enum CropAspectRatio
    {
        Free,
        Ratio1x1,
        Ratio4x3,
        Ratio16x9,
        OriginalRatio
    }

    public enum GradientType
    {
        Linear,
        Radial
    }
}

using System;

namespace OneAssist.Assistant.Core
{
    public enum RuleCategory
    {
        Formatting,
        Typography,
        Alignment,
        Colors,
        Layout,
        Editorial,
        Consistency,
        Accessibility,
        Images,
        Branding,
        Animation,
        FileHealth,
        ContentStructure
    }

    public enum IssueSeverity
    {
        Error,
        Warning,
        Info,
        Suggestion
    }

    public enum FixSafetyLevel
    {
        Safe,       // Non-destructive: Align, Match font, Replace approved color, Normalize spacing
        Review,     // Visual/layout change: Rewrite content, Change layout, Replace image
        Dangerous   // Destructive: Delete notes, Delete hidden slides, Remove links
    }

    public enum AppMode
    {
        BusinessUser, // Simplified: Scan, Fix Formatting, Apply Brand, Clean
        Designer,     // Full toolbox: Inspector, Layers, Align, Grid, Color, Morph
        Expert        // Advanced: Tokens, Rule editor, Diagnostics, Batch script
    }

    public enum FixActionType
    {
        AutoFix,
        ManualFix,
        GuidedFix
    }

    public enum AnalysisScope
    {
        ActiveSlide,
        SelectedShapes,
        SelectedSlides,
        EntirePresentation
    }

    public enum AlignmentAction
    {
        AlignLeft,
        AlignRight,
        AlignTop,
        AlignBottom,
        AlignHorizontalCenter,
        AlignVerticalCenter,
        DistributeHorizontally,
        DistributeVertically,
        CenterOnSlideHorizontally,
        CenterOnSlideVertically,
        MatchWidth,
        MatchHeight,
        MatchBothSize,
        EqualSpacingHorizontally,
        EqualSpacingVertically
    }
}

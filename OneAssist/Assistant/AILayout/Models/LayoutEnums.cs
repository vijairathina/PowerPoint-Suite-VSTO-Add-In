using System;

namespace OneAssist.Assistant.AILayout.Models
{
    public enum SlideContentType
    {
        TitleSlide,
        ExecutiveSummary,
        KPI,
        Comparison,
        Timeline,
        Process,
        Workflow,
        Strategy,
        Roadmap,
        Financial,
        DataAnalysis,
        Chart,
        Table,
        TextHeavy,
        ImageHeavy,
        BeforeAfter,
        Quote,
        CaseStudy,
        Architecture,
        OrgChart,
        Agenda,
        SectionDivider,
        Conclusion,
        GeneralContent
    }

    public enum ObjectSemanticType
    {
        // Text
        Title,
        Subtitle,
        Body,
        Label,
        Footnote,
        KPI,
        Description,
        Callout,

        // Images
        HeroImage,
        SupportingImage,
        Logo,
        Icon,

        // Charts
        PrimaryChart,
        SupportingChart,

        // Shapes & Containers
        Card,
        Container,
        Button,
        Decoration,
        Divider,

        // Tables
        Table,

        Unknown
    }

    public enum LayoutTemplatePattern
    {
        // Executive
        BigMessageEvidence,
        KPIDashboard,
        ExecutiveSummary,
        ThreeKeyTakeaways,
        OneBigNumber,

        // Comparison
        TwoColumnComparison,
        ThreeColumnComparison,
        BeforeAfter,
        ProsCons,

        // Process & Timeline
        HorizontalProcess,
        VerticalProcess,
        TimelineProcess,
        WorkflowJourney,

        // Strategy & Structure
        StrategyPillars,
        Matrix2x2,
        Roadmap,
        StrategicFramework,

        // Data & Charts
        ChartPlusInsight,
        KPIPlusChart,
        ChartPlusCommentary,
        DataDashboard,

        // Content & Cards
        TextPlusImage,
        ImagePlusText,
        ThreeCards,
        FourCards,
        QuotePlusImage,

        // Visual
        FullBleedHero,
        ImageGrid
    }

    public enum DesignStyle
    {
        Corporate,
        Modern,
        Minimal,
        Consulting,
        Marketing,
        Technical,
        Financial,
        Executive
    }

    public enum AIProviderType
    {
        LocalHeuristic,
        OpenAI,
        AzureOpenAI,
        Gemini,
        Ollama
    }
}

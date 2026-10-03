# PowerPoint VSTO – Slide Compare & Content Difference Module

## Overview

The **Slide Compare & Content Difference Module** is an isolated, enterprise-grade PowerPoint comparison engine designed for PowerPoint VSTO Add-ins (.NET Framework 4.8 / .NET Core / .NET 8).

Rather than relying on fragile slide indices or shape IDs, the module uses multi-dimensional similarity heuristics and content fingerprinting to compare presentations at multiple levels:
- **Presentation Level**: Structure, counts, added/deleted/reordered slides
- **Slide Level**: 1-to-1 matches, reorders, splits (1 slide → multiple slides), merges (multiple slides → 1 slide)
- **Shape Level**: Weighted 6-factor similarity (Text 40%, Type 15%, Position 15%, Size 10%, Name/Placeholder 10%, Visual 10%)
- **Content Level**: Word-level and character-level Longest Common Subsequence (LCS) text difference engine
- **Formatting Level**: Typography (font family, size, bold, italic, underline, color), alignment, geometry, dimensions, and rotation
- **Navigation**: Real-time bi-directional selection and navigation to destination shapes and slides in PowerPoint

---

## Architecture & Directory Structure

```text
SlideCompare/
│
├── Models/
│   ├── ComparisonEnums.cs          # SlideMatchType, ShapeMatchType, FilterMode, TextDiffType
│   ├── ShapeSnapshot.cs            # COM-free POCO shape model (text, geometry, format, groups, tables, charts)
│   ├── SlideSnapshot.cs            # COM-free POCO slide model with titles, notes, shapes, fingerprints
│   ├── PresentationSnapshot.cs     # POCO presentation model
│   └── ComparisonResult.cs         # Hierarchical results: Text diff tokens, shape diffs, slide diffs
│
├── Extraction/
│   ├── PresentationExtractor.cs    # Extracts PresentationSnapshot from PowerPoint COM with cancellation & progress
│   ├── SlideExtractor.cs           # Extracts SlideSnapshot, detecting titles, notes, and hidden slides
│   ├── ShapeExtractor.cs           # Deep recursive shape extractor (groups, tables, charts, pictures)
│   ├── TableExtractor.cs           # Cell-level table extractor
│   ├── ChartExtractor.cs           # Chart titles, series, and categories extractor
│   └── TextExtractor.cs            # Safe COM text frame and font property extraction
│
├── Matching/
│   ├── SimilarityCalculator.cs     # Text normalization, Levenshtein distance, Dice token similarity, geometric similarity
│   ├── TextDiffEngine.cs           # Word-level / Character-level LCS diff producing added, deleted, and modified tokens
│   ├── ContentMatchingEngine.cs    # SHA-256 semantic content hashing and fingerprinting
│   ├── ShapeMatchingEngine.cs      # Bipartite matching using 6-dimension weighted similarity
│   └── SlideMappingEngine.cs       # Slide mapping resolving 1-to-1, reorders, 1-to-many splits, and many-to-1 merges
│
├── Comparison/
│   ├── FormattingComparisonEngine.cs # Detects typography, color, size, rotation, and alignment changes
│   ├── SlideComparisonEngine.cs      # Slide comparison orchestrator
│   ├── ShapeComparisonEngine.cs      # Shape collection comparison
│   └── ComparisonEngine.cs           # Master presentation comparison pipeline
│
├── Navigation/
│   └── PowerPointNavigationService.cs # PowerPoint COM automation: activates destination slide & selects target shape
│
├── UI/
│   ├── SlideCompareControl.xaml      # Modern dark-theme WPF UI with summary pills, search, filter, diff boxes
│   ├── SlideCompareControl.xaml.cs   # Code-behind
│   ├── SlideCompareConverters.cs     # WPF status brushes and visibility converters
│   ├── SlideCompareViewModel.cs      # Full MVVM binding, async background execution, filters, and JSON/HTML export
│   ├── SlideCompareTaskPaneHost.cs   # WinForms UserControl container with ElementHost
│   └── SlideCompareTaskPaneManager.cs # Singleton VSTO CustomTaskPane lifecycle manager
│
├── Services/
│   ├── CompareSettings.cs            # Configurable weights, tolerances, and comparison toggles
│   ├── ComparisonService.cs          # Opens source presentation read-only in background, runs comparison async
│   └── SlideCompareManager.cs        # Public static entry point for one-line start & Ribbon integration
│
└── README.md                         # This integration guide
```

---

## Integration Guide: Copying into Another VSTO Project

The module has **zero third-party NuGet dependencies**! It relies exclusively on standard .NET Framework / .NET and Microsoft Office Interop assemblies.

### 1. Copy Files
Copy the entire `SlideCompare/` folder into your VSTO project directory.

### 2. Required References
Ensure the following references are present in your `.csproj`:
- `PresentationCore`
- `PresentationFramework`
- `System.Xaml`
- `WindowsBase`
- `WindowsFormsIntegration`
- `System.Windows.Forms`
- `System.Drawing`
- `System.Runtime.Serialization`
- `Microsoft.Office.Interop.PowerPoint`
- `Office` (`Microsoft.Office.Core`)
- `Microsoft.Office.Tools.Common`

### 3. Add to `.csproj`
In your `.csproj`, include the compile items and XAML page:
```xml
<ItemGroup>
  <Compile Include="SlideCompare\**\*.cs" />
  <Page Include="SlideCompare\UI\SlideCompareControl.xaml">
    <Generator>MSBuild:Compile</Generator>
    <SubType>Designer</SubType>
  </Page>
</ItemGroup>
```

### 4. Ribbon Command Integration
In your Ribbon designer (or XML ribbon):
1. Add a Ribbon Button labeled **Slide Compare** (icon: `CompareMergePresentations`).
2. In the click handler, add:
```csharp
private void btnSlideCompare_Click(object sender, RibbonControlEventArgs e)
{
    OneAssist.SlideCompare.Services.SlideCompareManager.Start();
}
```

### 5. Task Pane Cleanup on Shutdown
In `ThisAddIn.cs`:
```csharp
private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
{
    try
    {
        OneAssist.SlideCompare.UI.SlideCompareTaskPaneManager.Instance.Dispose();
    }
    catch { }
}
```

---

## Public API Usage Examples

### Programmatic Comparison
```csharp
using OneAssist.SlideCompare.Services;
using OneAssist.SlideCompare.Models;

// 1. Launch with default UI and file browser
SlideCompareManager.Start();

// 2. Launch with specific destination and source path
SlideCompareManager.Start(destinationPresentation, @"C:\Presentations\OriginalDeck.pptx");

// 3. Headless comparison without UI
var service = new ComparisonService();
var settings = new CompareSettings
{
    CompareFormatting = true,
    SimilarityThreshold = 0.75,
    PositionTolerancePoints = 12.0f
};

var result = await service.RunComparisonAsync(
    app,
    destinationPresentation,
    sourceFilePath,
    settings);

// Export results
string json = service.ExportToJson(result);
string html = service.ExportToHtml(result);
```

---

## Key Features & Comparison Philosophy

### 1. Intelligent Slide Mapping (`SlideMappingEngine`)
- **No hardcoded indices**: Never assumes Slide 1 == Slide 1.
- **Reordered slides**: Accurately detects when a slide is moved from position 2 to position 5 while keeping its contents intact.
- **Split Slides (1-to-many)**: Detects when a dense source slide is broken down across 2 or more destination slides.
- **Merged Slides (many-to-1)**: Detects when multiple source slides are condensed into 1 destination slide.

### 2. Shape-Level Matching (`ShapeMatchingEngine`)
Calculates a 6-dimension weighted similarity score:
```text
Text similarity:       40%
Shape type:            15%
Position similarity:   15%
Size similarity:       10%
Name/placeholder:      10%
Visual/content:        10%
```

### 3. Word-Level Text Difference (`TextDiffEngine`)
Uses Longest Common Subsequence to extract granular text modifications:
- Source: `"Revenue increased by 25% in 2025."`
- Destination: `"Revenue increased by 30% in 2025."`
- Display: `Revenue increased by [- 25%] [+ 30%] in 2025.`

### 4. Interactive Destination Navigation (`PowerPointNavigationService`)
- Clicking any result in the side panel automatically brings the destination slide into view and selects the target shape.
- For deleted items, navigates to the closest related destination slide and displays the deleted source content in the side panel.

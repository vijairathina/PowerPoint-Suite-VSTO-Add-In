using System;

namespace OneAssist.SlideCompare.Models
{
    /// <summary>
    /// Type of match determined for a slide comparison.
    /// </summary>
    public enum SlideMatchType
    {
        Matched,
        Modified,
        Added,
        Deleted,
        Moved,
        Reordered,
        Split,
        Merged
    }

    /// <summary>
    /// Type of match determined for a shape comparison.
    /// </summary>
    public enum ShapeMatchType
    {
        Matched,
        Modified,
        Added,
        Deleted,
        Moved
    }

    /// <summary>
    /// Category of difference found in a shape or slide.
    /// </summary>
    public enum DifferenceCategory
    {
        None,
        Content,
        Formatting,
        Geometry,
        Layout,
        Structure
    }

    /// <summary>
    /// Type of granular diff for text tokens.
    /// </summary>
    public enum TextDiffType
    {
        Unchanged,
        Added,
        Deleted,
        Modified,
        Moved
    }

    /// <summary>
    /// Filter mode for displaying comparison results.
    /// </summary>
    public enum FilterMode
    {
        IssuesOnly,
        All,
        Matched,
        Modified,
        Added,
        Deleted,
        Moved,
        Reordered,
        Split,
        Merged,
        TextChanges,
        ShapeChanges,
        FormattingChanges
    }

    /// <summary>
    /// Text comparison mode.
    /// </summary>
    public enum TextDiffMode
    {
        WordLevel,
        CharacterLevel
    }
}

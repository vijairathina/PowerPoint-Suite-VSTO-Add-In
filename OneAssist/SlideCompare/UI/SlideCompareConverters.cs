using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using OneAssist.SlideCompare.Models;

namespace OneAssist.SlideCompare.UI
{
    public class BoolToVisibilityConverter : IValueConverter
    {
        public bool Invert { get; set; }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool val = value is bool b && b;
            if (Invert) val = !val;
            return val ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class SlideStatusColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SlideMatchType matchType)
            {
                switch (matchType)
                {
                    case SlideMatchType.Matched:
                        return new SolidColorBrush(Color.FromRgb(34, 197, 94)); // Emerald Green #22C55E
                    case SlideMatchType.Modified:
                        return new SolidColorBrush(Color.FromRgb(245, 158, 11)); // Amber #F59E0B
                    case SlideMatchType.Added:
                        return new SolidColorBrush(Color.FromRgb(56, 189, 248)); // Sky Blue #38BDF8
                    case SlideMatchType.Deleted:
                        return new SolidColorBrush(Color.FromRgb(239, 68, 68)); // Crimson Red #EF4444
                    case SlideMatchType.Moved:
                    case SlideMatchType.Reordered:
                        return new SolidColorBrush(Color.FromRgb(251, 191, 36)); // Yellow #FBBF24
                    case SlideMatchType.Split:
                        return new SolidColorBrush(Color.FromRgb(6, 182, 212)); // Cyan Teal #06B6D4
                    case SlideMatchType.Merged:
                        return new SolidColorBrush(Color.FromRgb(236, 72, 153)); // Magenta Pink #EC4899
                }
            }
            return new SolidColorBrush(Color.FromRgb(100, 116, 139));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class ShapeStatusColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ShapeMatchType matchType)
            {
                switch (matchType)
                {
                    case ShapeMatchType.Matched:
                        return new SolidColorBrush(Color.FromRgb(34, 197, 94)); // Green #22C55E
                    case ShapeMatchType.Modified:
                        return new SolidColorBrush(Color.FromRgb(245, 158, 11)); // Amber #F59E0B
                    case ShapeMatchType.Added:
                        return new SolidColorBrush(Color.FromRgb(56, 189, 248)); // Blue #38BDF8
                    case ShapeMatchType.Deleted:
                        return new SolidColorBrush(Color.FromRgb(239, 68, 68)); // Red #EF4444
                    case ShapeMatchType.Moved:
                        return new SolidColorBrush(Color.FromRgb(251, 191, 36)); // Yellow #FBBF24
                }
            }
            return new SolidColorBrush(Color.FromRgb(100, 116, 139));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Attached property to render rich diff tokens inside a WPF TextBlock.
    /// Deleted text: Red (#EF4444) with Strikethrough
    /// Inserted text: Blue (#38BDF8) bold
    /// Moved text: Yellow (#FBBF24) semi-bold
    /// Unchanged text: Light Slate (#CBD5E1)
    /// </summary>
    public static class DiffTextBlockHelper
    {
        public static readonly DependencyProperty DiffTokensProperty =
            DependencyProperty.RegisterAttached(
                "DiffTokens",
                typeof(IEnumerable<TextDiffToken>),
                typeof(DiffTextBlockHelper),
                new PropertyMetadata(null, OnDiffTokensChanged));

        public static IEnumerable<TextDiffToken> GetDiffTokens(DependencyObject obj) =>
            (IEnumerable<TextDiffToken>)obj.GetValue(DiffTokensProperty);

        public static void SetDiffTokens(DependencyObject obj, IEnumerable<TextDiffToken> value) =>
            obj.SetValue(DiffTokensProperty, value);

        private static void OnDiffTokensChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TextBlock textBlock)
            {
                textBlock.Inlines.Clear();
                if (e.NewValue is IEnumerable<TextDiffToken> tokens)
                {
                    var redBrush = new SolidColorBrush(Color.FromRgb(239, 68, 68)); // Red #EF4444
                    var blueBrush = new SolidColorBrush(Color.FromRgb(56, 189, 248)); // Blue #38BDF8
                    var yellowBrush = new SolidColorBrush(Color.FromRgb(251, 191, 36)); // Yellow #FBBF24
                    var normalBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225)); // Slate #CBD5E1

                    redBrush.Freeze();
                    blueBrush.Freeze();
                    yellowBrush.Freeze();
                    normalBrush.Freeze();

                    foreach (var token in tokens)
                    {
                        if (string.IsNullOrEmpty(token.Text)) continue;

                        switch (token.DiffType)
                        {
                            case TextDiffType.Deleted:
                                textBlock.Inlines.Add(new Run(token.Text)
                                {
                                    Foreground = redBrush,
                                    TextDecorations = TextDecorations.Strikethrough,
                                    FontWeight = FontWeights.Medium
                                });
                                break;

                            case TextDiffType.Added:
                                textBlock.Inlines.Add(new Run(token.Text)
                                {
                                    Foreground = blueBrush,
                                    FontWeight = FontWeights.Bold
                                });
                                break;

                            case TextDiffType.Moved:
                            case TextDiffType.Modified:
                                textBlock.Inlines.Add(new Run(token.Text)
                                {
                                    Foreground = yellowBrush,
                                    FontWeight = FontWeights.SemiBold
                                });
                                break;

                            case TextDiffType.Unchanged:
                            default:
                                textBlock.Inlines.Add(new Run(token.Text)
                                {
                                    Foreground = normalBrush
                                });
                                break;
                        }
                    }
                }
            }
        }
    }
}

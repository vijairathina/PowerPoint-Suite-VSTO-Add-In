using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media;
using OneAssist.Assistant.Core;

namespace OneAssist.Assistant.UI
{
    public class SeverityToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is IssueSeverity severity)
            {
                switch (severity)
                {
                    case IssueSeverity.Error:
                        return new SolidColorBrush(Color.FromRgb(220, 53, 69)); // Crimson Red
                    case IssueSeverity.Warning:
                        return new SolidColorBrush(Color.FromRgb(235, 130, 20)); // Amber Orange
                    case IssueSeverity.Info:
                        return new SolidColorBrush(Color.FromRgb(23, 162, 184)); // Cyan Blue
                    case IssueSeverity.Suggestion:
                        return new SolidColorBrush(Color.FromRgb(40, 167, 69)); // Emerald Green
                }
            }
            return new SolidColorBrush(Colors.Gray);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class HealthScoreToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double score)
            {
                if (score >= 90) return new SolidColorBrush(Color.FromRgb(40, 167, 69));
                if (score >= 70) return new SolidColorBrush(Color.FromRgb(235, 130, 20));
                return new SolidColorBrush(Color.FromRgb(220, 53, 69));
            }
            return new SolidColorBrush(Color.FromRgb(0, 122, 255));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class IntToBoolConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int intVal && int.TryParse(parameter?.ToString(), out int targetVal))
            {
                return intVal == targetVal;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b && b && int.TryParse(parameter?.ToString(), out int targetVal))
            {
                return targetVal;
            }
            return Binding.DoNothing;
        }

        public override object ProvideValue(IServiceProvider serviceProvider) => this;
    }

    public class TabIndexToVisConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int intVal && int.TryParse(parameter?.ToString(), out int targetVal))
            {
                return intVal == targetVal ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider) => this;
    }

    public class ModeToVisibilityConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && parameter != null)
            {
                string current = value.ToString();
                string target = parameter.ToString();
                bool negate = false;
                if (target.StartsWith("!"))
                {
                    negate = true;
                    target = target.Substring(1);
                }
                bool match = string.Equals(current, target, StringComparison.OrdinalIgnoreCase);
                if (negate) match = !match;
                return match ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public override object ProvideValue(IServiceProvider serviceProvider) => this;
    }

    public class ModeToBoolConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && parameter != null)
            {
                return string.Equals(value.ToString(), parameter.ToString(), StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool b && b && parameter != null)
            {
                if (Enum.TryParse(parameter.ToString(), out AppMode mode))
                {
                    return mode;
                }
            }
            return Binding.DoNothing;
        }

        public override object ProvideValue(IServiceProvider serviceProvider) => this;
    }

    public class NullToVisConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value == null ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
        public override object ProvideValue(IServiceProvider serviceProvider) => this;
    }

    public class NotNullToVisConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != null ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
        public override object ProvideValue(IServiceProvider serviceProvider) => this;
    }
}

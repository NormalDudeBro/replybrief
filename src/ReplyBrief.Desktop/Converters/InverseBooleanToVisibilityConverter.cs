using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ReplyBrief.Desktop.Converters;

/// <summary>
/// Shows an element when the bound flag is false, so a placeholder can sit in the same slot as
/// the real content without both being visible at once.
/// </summary>
internal sealed class InverseBooleanToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is true ? Visibility.Collapsed : Visibility.Visible;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
        value is Visibility.Visible ? false : true;
}

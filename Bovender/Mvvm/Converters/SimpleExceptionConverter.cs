using System;
using System.Windows;
using System.Windows.Data;

namespace Bovender.Mvvm.Converters
{
    /// <summary>
    /// WPF converter that produces a simple text string from an exception.
    /// </summary>
    /// <remarks>
    /// The built-in conversion of the WPF uses the ToString() method of
    /// the exceptions, which creates unwieldy long strings. This converter
    /// simply returns the normal human-friendly exception message.
    /// </remarks>
    public class SimpleExceptionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            Exception e = value as Exception;
            if (e != null)
            {
                return e.Message;
            }
            else
            {
                return DependencyProperty.UnsetValue;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return DependencyProperty.UnsetValue;
        }
    }
}

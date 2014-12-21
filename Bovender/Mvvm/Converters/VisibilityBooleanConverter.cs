using System;
using System.Windows;
using System.Windows.Data;

namespace Bovender.Mvvm.Converters
{
    public class VisibilityBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((bool)value == true)
            {
                return Visibility.Visible;
            }
            else
            {
                return Visibility.Collapsed;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            switch ((Visibility)value)
            {
                case Visibility.Visible: return true;
                case Visibility.Hidden: return false;
                case Visibility.Collapsed: return false;
                default:
                    throw new ArgumentException(
                        "No conversion defined for " + ((Visibility)value).ToString());
            }
        }
    }
}

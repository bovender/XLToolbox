using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;

namespace Bovender.Mvvm.Converters
{
    /// <summary>
    /// Converts enum values to booleans to enable easy use
    /// of enumerations with WPF radio buttons.
    /// </summary>
    /// <remarks>
    /// Credits to Scott @ http://stackoverflow.com/a/2908885/270712
    /// </remarks>
    /// <example><code><![CDATA[
    ///     <StackPanel>
    ///         <StackPanel.Resources>          
    ///             <local:EnumToBooleanConverter x:Key="ebc" />          
    ///         </StackPanel.Resources>
    ///         <RadioButton IsChecked="{Binding Path=YourEnumProperty,
    ///                      Converter={StaticResource ebc},
    ///                      ConverterParameter={x:Static local:YourEnumType.Enum1}}" />
    ///         <RadioButton IsChecked="{Binding Path=YourEnumProperty,
    ///                      Converter={StaticResource ebc},
    ///                      ConverterParameter={x:Static local:YourEnumType.Enum2}}" />
    ///     </StackPanel>    /// <Grid>
    /// ]]></code></example>
    class EnumBooleanConverter : IValueConverter
    {
        #region IValueConverter interface

        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            return value.Equals(parameter);
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            return value.Equals(true) ? parameter : Binding.DoNothing;
        }

        #endregion
    }
}

/* BooleanNegationConverter.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Data;

namespace Bovender.Mvvm.Converters
{
    /// <summary>
    /// Negates a boolean value.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Credits: http://stackoverflow.com/a/1039681/270712
    /// </para>
    /// <para>
    /// For ease of use, put something like this in a central resource dictionary:
    /// <code>
    ///     <ResourceDictionary  
    ///         xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    ///         xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    ///         xmlns:conv="clr-namespace:Bovender.Mvvm.Converters;assembly=Bovender"
    ///         >
    ///         <conv:BooleanNegationConverter x:Key="boolNegConv" />
    ///     </ResourceDictionary>
    /// </code>
    /// </para>
    /// </remarks>
    [ValueConversion(typeof(bool), typeof(bool))]
    public class BooleanNegationConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool))
                throw new InvalidOperationException("The target must be a boolean");

            return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        #endregion
    }
}

/* SimpleExceptionConverter.cs
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

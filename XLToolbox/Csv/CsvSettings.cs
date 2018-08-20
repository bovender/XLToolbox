/* CsvSettings.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using System.Globalization;
using System.Linq;
using System.Text;

namespace XLToolbox.Csv
{
    public class CsvSettings
    {
        #region Public properties

        public string FieldSeparator { get; set; }

        public string DecimalSeparator
        {
            get { return NumberFormatInfo.NumberDecimalSeparator; }
            set { NumberFormatInfo.NumberDecimalSeparator = value; }
        }

        public string ThousandsSeparator
        {
            get { return NumberFormatInfo.NumberGroupSeparator; }
            set { NumberFormatInfo.NumberGroupSeparator = value; }
        }

        public bool Tabularize { get; set; }

        /// <summary>
        /// Returns a System.Globalization.NumberFormatInfo object
        /// whose properties are set according to the current
        /// properties (namely, <see cref="DecimalSeparator"/>.)
        /// This property is mainly used internally, but available
        /// publicly for convenience.
        /// </summary>
        [YamlDotNet.Serialization.YamlIgnore]
        public NumberFormatInfo NumberFormatInfo
        {
            get
            {
                if (_numberFormatInfo == null)
                {
                    _numberFormatInfo = CultureInfo.InvariantCulture.NumberFormat.Clone() as NumberFormatInfo;
                    _numberFormatInfo.NumberGroupSeparator = "";
                }
                return _numberFormatInfo;
            }
        }

        #endregion

        #region Constructor

        public CsvSettings()
        {
            FieldSeparator = ",";
            DecimalSeparator = ".";
            ThousandsSeparator = "";
        }

        #endregion

        #region Fields

        private NumberFormatInfo _numberFormatInfo;

        #endregion
    }
}

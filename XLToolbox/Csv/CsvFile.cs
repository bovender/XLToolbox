/* CsvFile.cs
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

namespace XLToolbox.Csv
{
    /// <summary>
    /// Provides import/export settings and methods for CSV files.
    /// </summary>
    [Serializable]
    public class CsvFile : Object
    {
        #region Factory

        public static CsvFile FromLastUsed()
        {
            CsvFile c = Properties.Settings.Default.CsvSettings;
            if (c == null)
            {
                c = new CsvFile();
            }
            return c;
        }

        #endregion

        #region Public properties

        public string FileName { get; set; }

        public string FieldSeparator { get; set; }

        public string DecimalSeparator { get; set; }

        public string ThousandsSeparator { get; set; }

        #endregion

        #region Constructor

        public CsvFile()
        {
            FieldSeparator = ",";
            DecimalSeparator = ".";
            ThousandsSeparator = ",";
        }

        #endregion

        #region Import/export methods

        public void Import()
        {
            Properties.Settings.Default.CsvSettings = this;
            Properties.Settings.Default.Save();
            Excel.ViewModels.Instance.Default.Application.Workbooks.OpenText(
                FileName,
                DataType: Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited,
                Other: true, OtherChar: StringParam(FieldSeparator),
                DecimalSeparator: StringParam(DecimalSeparator),
                ThousandsSeparator: StringParam(ThousandsSeparator),
                Local: false, ConsecutiveDelimiter: false,
                Origin: Microsoft.Office.Interop.Excel.XlPlatform.xlWindows
                );
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Helper function that converts empty strings to Type.Missing.
        /// </summary>
        /// <param name="s">String to convert.</param>
        /// <returns>String or Type.Missing if string is null or empty.</returns>
        /// <remarks>This function is necessary because Workbooks.OpenText will
        /// throw a COM exception if one of the optional parameters is an empty
        /// string.
        /// </remarks>
        object StringParam(string s)
        {
            if (String.IsNullOrEmpty(s))
            {
                return Type.Missing;
            }
            else
            {
                return s;
            }
        }

        #endregion
    }
}

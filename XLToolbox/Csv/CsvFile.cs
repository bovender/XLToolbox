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
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace XLToolbox.Csv
{
    /// <summary>
    /// Provides import/export settings and methods for CSV files.
    /// </summary>
    [Serializable]
    public class CsvFile : Object
    {
        #region Factory

        public static CsvFile LastImport()
        {
            CsvFile c = Properties.Settings.Default.CsvImport;
            if (c == null)
            {
                c = Properties.Settings.Default.CsvExport;
                if (c == null)
                {
                    c = new CsvFile();
                }
            }
            return c;
        }

        public static CsvFile LastExport()
        {
            CsvFile c = Properties.Settings.Default.CsvExport;
            if (c == null)
            {
                c = Properties.Settings.Default.CsvImport;
                if (c == null)
                {
                    c = new CsvFile();
                }
            }
            return c;
        }

        #endregion

        #region Public properties

        public string FileName { get; set; }

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

        /// <summary>
        /// Returns a System.Globalization.NumberFormatInfo object
        /// whose properties are set according to the current
        /// properties (namely, <see cref="DecimalSeparator"/>.)
        /// This property is mainly used internally, but available
        /// publicly for convenience.
        /// </summary>
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
            Properties.Settings.Default.CsvImport = this;
            Properties.Settings.Default.Save();
            Excel.ViewModels.Instance.Default.Application.Workbooks.OpenText(
                FileName,
                DataType: XlTextParsingType.xlDelimited,
                Other: true, OtherChar: StringParam(FieldSeparator),
                DecimalSeparator: StringParam(DecimalSeparator),
                ThousandsSeparator: StringParam(ThousandsSeparator),
                Local: false, ConsecutiveDelimiter: false,
                Origin: XlPlatform.xlWindows
                );
        }

        /// <summary>
        /// Exports the used range of the active worksheet to a CSV file,
        /// using the file name specified in the <see cref="FileName"/> property.
        /// </summary>
        public void Export()
        {
            Worksheet worksheet = Excel.ViewModels.Instance.Default.ActiveWorkbook.ActiveSheet as Worksheet;
            if (worksheet == null)
            {
                throw new InvalidOperationException("Cannot export chart to CSV file.");
            }
            Export(worksheet.UsedRange);
        }

        /// <summary>
        /// Exports a range to a CSV file, using the file name specified
        /// in the <see cref="FileName"/> property.
        /// </summary>
        /// <param name="range">Range to export.</param>
        public void Export(Range range)
        {
            Properties.Settings.Default.CsvExport = this;
            Properties.Settings.Default.Save();
            StreamWriter sw = File.CreateText(FileName);
            foreach (Range row in range.Rows)
            {
                bool needSep = false;
                foreach (Range cell in row.Cells)
                {
                    // If this is not the first field in the line, write a field separator.
                    if (needSep)
                    {
                        sw.Write(FieldSeparator);
                    }
                    else
                    {
                        needSep = true;
                    }

                    // Range.Value2 is declared as dynamic and may return
                    // a string or a double. We let the runtime take care
                    // of this by calling an overloaded method with the dynamic.
                    sw.Write(FieldToStr(cell.Value2));
                }
                sw.WriteLine();
            }
            sw.Close();
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Overloaded method to facilitate code switching base
        /// on the actual type of a dynamic variable as returned
        /// by Range.Value2 (which may be a string, a double, a ...).
        /// </summary>
        /// <param name="d">Double to convert to a string</param>
        /// <returns>String that can be written to a CSV file.</returns>
        string FieldToStr(double d)
        {
            return d.ToString(NumberFormatInfo);
        }

        /// <summary>
        /// Overloaded method to facilitate code switching base
        /// on the actual type of a dynamic variable as returned
        /// by Range.Value2 (which may be a string, a double, a ...).
        /// </summary>
        /// <param name="s">String field value.</param>
        /// <returns>String that can be written to a CSV file;
        /// if <paramref name="s"/> contains the current field
        /// separator, it will be wrapped in double quotes.</returns>
        string FieldToStr(string s)
        {
            if (s.Contains(FieldSeparator))
            {
                return "\"" + s + "\"";
            }
            else
	        {
                return s;
	        }
        }

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

        #region Fields

        NumberFormatInfo _numberFormatInfo;

        #endregion
    }
}

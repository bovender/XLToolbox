/* CsvFile.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2016 Daniel Kraus
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
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using XLToolbox.Excel;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace XLToolbox.Csv
{
    /// <summary>
    /// Provides import/export settings and methods for CSV files.
    /// </summary>
    public class CsvFile : Bovender.Mvvm.Models.ProcessModel
    {
        #region Factory

        public static CsvFile LastImport()
        {
            CsvFile c = UserSettings.UserSettings.Default.CsvImport;
            if (c == null)
            {
                c = new CsvFile();
            }
            return c;
        }

        public static CsvFile LastExport()
        {
            CsvFile c = UserSettings.UserSettings.Default.CsvExport;
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

        /// <summary>
        /// Gets whether the exporter is currently processing.
        /// </summary>
        [XmlIgnore]
        public bool IsProcessing { get; private set; }

        /// <summary>
        /// Gets the number of cells that were already processed
        /// during export.
        /// </summary>
        [XmlIgnore]
        public long CellsProcessed { get; private set; }

        /// <summary>
        /// Gets the total number of cells to export.
        /// </summary>
        [XmlIgnore]
        public long CellsTotal { get; private set; }

        #endregion

        #region Events

        public event EventHandler<ErrorEventArgs> ExportFailed;

        public event EventHandler<EventArgs> ExportProgressCompleted;

        #endregion

        #region Constructor

        public CsvFile()
        {
            FieldSeparator = ",";
            DecimalSeparator = ".";
            ThousandsSeparator = "";
        }

        #endregion

        #region Import/export methods

        public void Import()
        {
            Logger.Info("Importing CSV: FS='{0}', DS='{1}', TS='{2}'",
                FieldSeparator, DecimalSeparator, ThousandsSeparator);
            UserSettings.UserSettings.Default.CsvImport = this;
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
            Logger.Info("Exporting CSV: FS='{0}', DS='{1}', TS='{2}'",
                FieldSeparator, DecimalSeparator, ThousandsSeparator);
            UserSettings.UserSettings.Default.CsvExport = this;
            IsProcessing = true;
            Task t = new Task(() =>
            {
                StreamWriter sw = null;
                try
                {
                    // StreamWriter buffers the output; using a StringBuilder
                    // doesn't speed up things (tried it)
                    sw = File.CreateText(FileName);
                    CellsTotal = range.CellsCount();
                    Logger.Info("Number of cells: {0}", CellsTotal);
                    CellsProcessed = 0;
                    _cancelExport = false;
                    string fs = FieldSeparator;
                    if (fs == "\\t") { fs = "\t"; } // Convert "\t" to tab characters


                    // Get all values in an array
                    foreach (Range row in range.Rows)
                    {
                        // object[,] values = range.Value2;
                        object[,] values = row.Value2;
                        if (values != null)
                        {
                            //for (long row = 1; row <= values.GetLength(0); row++)
                            //{
                            for (long col = 1; col <= values.GetLength(1); col++)
                            {
                                CellsProcessed++;

                                // If this is not the first field in the line, write a field separator.
                                if (col > 1)
                                {
                                    sw.Write(fs);
                                }

                                // object value = values[row, col];
                                object value = values[1, col]; // 1-based index!
                                if (value != null)
                                {
                                    if (value is string)
                                    {
                                        string s = value as string;
                                        if (s.Contains(fs) || s.Contains("\""))
                                        {
                                            s = "\"" + s.Replace("\"", "\"\"") + "\"";
                                        }
                                        sw.Write(s);
                                    }
                                    else
                                    {
                                        double d = Convert.ToDouble(value);
                                        sw.Write(d.ToString(NumberFormatInfo));
                                    }
                                }
                                if (_cancelExport) break;
                            }
                            sw.WriteLine();
                        }
                        if (_cancelExport)
                        {
                            sw.WriteLine(UNFINISHED_EXPORT);
                            sw.WriteLine("Cancelled by user.");
                            Logger.Info("CSV export cancelled by user");
                            break;
                        }
                        // }
                    }
                    sw.Close();
                    if (!_cancelExport) OnProcessSucceeded();
                }
                catch (IOException e)
                {
                    IsProcessing = false;
                    OnProcessFailed(e);
                }
                catch (Exception e1)
                {
                    IsProcessing = false;
                    if (sw != null)
                    {
                        sw.WriteLine(UNFINISHED_EXPORT);
                        sw.WriteLine(e1.ToString());
                        sw.Close();
                    }
                    OnProcessFailed(e1);
                }
                finally
                {
                    IsProcessing = false;
                    Logger.Info("CSV export task finished");
                }
            });
            t.Start();
        }

        public void CancelExport()
        {
            _cancelExport = true;
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

        #region Fields

        NumberFormatInfo _numberFormatInfo;
        bool _cancelExport;

        #endregion

        #region Private constant

        const string UNFINISHED_EXPORT = "*** UNFINISHED EXPORT ***";
        #endregion
    }
}

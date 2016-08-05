/* CsvExporter.cs
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
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using YamlDotNet.Serialization;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Excel;
using System.Runtime.InteropServices;

namespace XLToolbox.Csv
{
    public class CsvExporter : CsvFileBase
    {
        #region Factory

        public static CsvExporter LastExport()
        {
            CsvSettings settings = UserSettings.UserSettings.Default.CsvSettings;
            if (settings == null)
            {
                return new CsvExporter();
            }
            else
            {
                return new CsvExporter(settings);
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets whether the exporter is currently processing.
        /// </summary>
        [YamlIgnore]
        public bool IsProcessing { get; private set; }

        /// <summary>
        /// Gets the number of cells that were already processed
        /// during export.
        /// </summary>
        [YamlIgnore]
        public long CellsProcessed { get; private set; }

        /// <summary>
        /// Gets the total number of cells to export.
        /// </summary>
        [YamlIgnore]
        public long CellsTotal { get; private set; }

        [YamlIgnore]
        public Range Range { get; set; }
        
        #endregion

        #region Constructor

        public CsvExporter()
            : base()
        { }

        public CsvExporter(CsvSettings settings)
            : base(settings)
        { }

        #endregion

        #region Implementation of ProcessModel

        /// <summary>
        /// Exports the Range to a CSV file, using the file name specified
        /// in the <see cref="FileName"/> property.
        /// </summary>
        /// <param name="range">Range to export.</param>
        public override bool Execute()
        {
            Logger.Info("Export: Exporting CSV: FS='{0}', DS='{1}', TS='{2}'",
                FieldSeparator, DecimalSeparator, ThousandsSeparator);
            bool result = false;
            UserSettings.UserSettings.Default.CsvSettings = Settings;
            if (Range == null)
            {
                Range = UsedRange();
            }
            if (Range == null)
            {
                throw new InvalidOperationException("Cannot export CSV: No range given, and the used range cannot be determined.");
            }

            StreamWriter sw = null;
            try
            {
                // StreamWriter buffers the output; using a StringBuilder
                // doesn't speed up things (tried it)
                sw = File.CreateText(FileName);
                CellsTotal = Range.CellsCount();
                Logger.Info("Number of cells: {0}", CellsTotal);
                CellsProcessed = 0;
                string fs = FieldSeparator;
                if (fs == "\\t") { fs = "\t"; } // Convert "\t" to tab characters

                // Get all values in an array
                Range rows = Range.Rows;
                foreach (Range row in rows)
                {
                    // object[,] values = range.Value2;
                    object[,] values = row.Value2;
                    if (values != null)
                    {
                        for (long col = 1; col <= values.GetLength(1); col++)
                        {
                            CellsProcessed++;

                            // If this is not the first field in the line, write a field separator.
                            if (col > 1)
                            {
                                sw.Write(fs);
                            }

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
                            if (IsCancellationRequested) break;
                        }
                        sw.WriteLine();
                        if (Marshal.IsComObject(values)) Marshal.ReleaseComObject(values);
                    }
                    if (IsCancellationRequested)
                    {
                        sw.WriteLine(UNFINISHED_EXPORT);
                        sw.WriteLine("Cancelled by user.");
                        Logger.Info("CSV export cancelled by user");
                        break;
                    }
                    if (Marshal.IsComObject(row)) Marshal.ReleaseComObject(row);
                }
                if (Marshal.IsComObject(rows)) Marshal.ReleaseComObject(rows);
                sw.Close();
            }
            catch (IOException ioException)
            {
                IsProcessing = false;
                throw ioException;
            }
            catch (Exception anyException)
            {
                IsProcessing = false;
                if (sw != null)
                {
                    sw.WriteLine(UNFINISHED_EXPORT);
                    sw.WriteLine(anyException.ToString());
                    sw.Close();
                }
                throw anyException;
            }
            finally
            {
                IsProcessing = false;
                Logger.Info("CSV export task finished");
            }
            return result;
        }

        #endregion

        #region Private helper methods

        private Range UsedRange()
        {
            Worksheet worksheet = Excel.ViewModels.Instance.Default.ActiveWorkbook.ActiveSheet as Worksheet;
            if (worksheet != null)
            {
                return worksheet.UsedRange;
            }
            else
            {
                return null;
            }
        }

        #endregion

        #region Private constant

        const string UNFINISHED_EXPORT = "*** UNFINISHED EXPORT ***";

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

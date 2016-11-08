using Microsoft.Office.Interop.Excel;
/* CsvImporter.cs
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
using System.Linq;
using System.Text;
using System.IO;
using Bovender.Extensions;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Csv
{
    public class CsvImporter : CsvFileBase
    {
        #region Factory

        public static CsvImporter LastImport()
        {
            CsvSettings settings = UserSettings.UserSettings.Default.CsvSettings;
            if (settings == null)
            {
                return new CsvImporter();
            }
            else
            {
                return new CsvImporter(settings);
            }
        }

        #endregion

        #region Implementation of ProcessModel

        /// <summary>
        /// Opens a text file 
        /// </summary>
        /// <returns></returns>
        public override bool Execute()
        {
            bool result;

            Logger.Info("Importing CSV: FS='{0}', DS='{1}', TS='{2}'",
                FieldSeparator, DecimalSeparator, ThousandsSeparator);
            UserSettings.UserSettings.Default.CsvSettings = Settings;

            if (Path.GetExtension(FileName).ToUpper() == ".TXT")
            {
                Logger.Info("Execute: .txt file, opening directly");
                OpenText(FileName);
                result = true;
            }
            else
            {
                Logger.Info("Execute: Not a .txt file, creating temporary copy");
                string tempName = Path.ChangeExtension(Path.GetTempFileName(), ".txt");
                Logger.Debug("Execute: {0}", tempName);
                File.Copy(FileName, tempName, true);
                Instance.Default.DisableScreenUpdating();
                try
                {
                    OpenText(tempName);
                    Workbook tempWb = Instance.Default.ActiveWorkbook;
                    Worksheet sheet = tempWb.ActiveSheet as Worksheet;
                    Logger.Info("Execute: Copying sheet");
                    sheet.Name = Path.GetFileName(FileName).TruncateWithEllipsis(31);
                    tempWb.Saved = true;
                    sheet.Copy();
                    tempWb.Close();
                    Bovender.ComHelpers.ReleaseComObject(sheet);
                    Bovender.ComHelpers.ReleaseComObject(tempWb);
                    Logger.Info("Execute: Deleting temporary file");
                    File.Delete(tempName);
                    result = true;
                }
                catch (Exception e)
                {
                    Logger.Warn("Execute: Caught exception");
                    Logger.Warn(e);
                    throw e;
                }
                finally
                {
                    Instance.Default.EnableScreenUpdating();
                }
            }

            return result;
        }
        
        #endregion

        #region Constructors

        public CsvImporter()
            : base()
        { }

        public CsvImporter(CsvSettings settings)
            : base(settings)
        { }

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

        private void OpenText(string fileName)
        {
            Excel.ViewModels.Instance.Default.Workbooks.OpenText(
                fileName,
                DataType: XlTextParsingType.xlDelimited,
                Other: true, OtherChar: StringParam(FieldSeparator),
                DecimalSeparator: StringParam(DecimalSeparator),
                ThousandsSeparator: StringParam(ThousandsSeparator),
                Local: false, ConsecutiveDelimiter: false,
                Origin: XlPlatform.xlWindows
                );
        }

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

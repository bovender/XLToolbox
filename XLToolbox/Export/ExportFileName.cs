/* ExportFileName.cs
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
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Export.Models;

namespace XLToolbox.Export
{
    /// <summary>
    /// Generates file names to use with for graphics export.
    /// </summary>
    public class ExportFileName
    {
        #region Public properties

        public int Counter { get; protected set; }

        public string Directory { get; private set; }

        #endregion

        #region Constructors

        public ExportFileName(string template, FileType fileType)
        {
            Template = template;
            Counter = 0;
            FileType = fileType;
            SetExtension();
            _placeholderReplacements = new Dictionary<string, Func<string>>()
            {
                { Strings.Workbook.ToUpper(), () =>
                    { return Path.GetFileNameWithoutExtension(this.CurrentWorkbookName); } },
                { Strings.Worksheet.ToUpper(), () => { return this.CurrentWorksheetName; } },
                { Strings.Index.ToUpper(), () => { return String.Format("{0:000}", Counter); } },
            };
            Directory = String.Empty;
        }

        public ExportFileName(string directory, string template, FileType fileType)
            : this(template, fileType)
        {
            Directory = directory;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Generates the next file name to use; this will increase the
        /// internal counter.
        /// </summary>
        /// <returns></returns>
        public string GenerateNext(dynamic worksheet)
        {
            CurrentWorkbookName = worksheet.Parent.Name;
            CurrentWorksheetName = worksheet.Name;
            Counter++;
            string s = _regex.Replace(Template, SubstituteVariable);
            // If no index placeholder exists in the template, add the index at the end.
            return Path.Combine(Directory, InsertIndexIfMissing(Template, s) + _extension);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Replaces a placeholder with the appropriate value, or returns
        /// the matched string unchanged.
        /// </summary>
        /// <param name="match">Placeholder match ("{..}").</param>
        /// <returns>Replacement string or match itself if no placeholder found.</returns>
        private string SubstituteVariable(Match match)
        {
            Func<string> func;
            // Cut leading and trailing {}, convert to upper case
            string placeholder = match.ToString().Substring(1, match.Value.Length - 2).ToUpper();
            if (_placeholderReplacements.TryGetValue(placeholder, out func) == true)
            {
                return func();
            }
            else
            {
                return match.Value;
            }
        }

        private string InsertIndexIfMissing(string template, string fileName)
        {
            if (!template.Contains("{" + Strings.Index + "}"))
            {
                return Path.GetFileNameWithoutExtension(fileName) +
                    String.Format("{0:000}", Counter) +
                    Path.GetExtension(fileName);
            }
            else
            {
                return fileName;
            }
        }

        private void SetExtension()
        {
            if (String.IsNullOrWhiteSpace(Template) || !Template.ToUpper().EndsWith(FileType.ToFileNameExtension().ToUpper()))
            {
                _extension = FileType.ToFileNameExtension();
            }
            else
            {
                _extension = String.Empty;
            }
        }

        #endregion

        #region Protected properties

        protected string Template { get; private set; }
        protected string CurrentWorkbookName { get; set; }
        protected string CurrentWorksheetName { get; set; }
        protected FileType FileType { get; private set; }

        #endregion

        #region Private fields

        Dictionary<string, Func<string>> _placeholderReplacements;
        string _extension;
        private static readonly Regex _regex = new Regex(@"{[^}]+}");

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

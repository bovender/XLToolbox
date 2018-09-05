/* ExportFileName.cs
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

        // TODO: Clean up those constructor signatures
        public ExportFileName(BatchExportSettings settings)
            : this(settings.Path, settings.FileName, settings.Preset.FileType, settings)
        { }

        public ExportFileName(string template, FileType fileType, BatchExportSettings settings)
        {
            Template = template;
            Counter = 0;
            FileType = fileType;
            _settings = settings;
            _placeholderReplacements = new Dictionary<string, Func<string>>()
            {
                { Strings.Workbook.ToUpper(), () =>
                    { return Path.GetFileNameWithoutExtension(this.CurrentWorkbookName); } },
                { Strings.Worksheet.ToUpper(), () => { return this.CurrentWorksheetName; } },
                { Strings.Index.ToUpper(), () => { return String.Format("{0:000}", Counter); } },
                { Strings.Name.ToUpper(), () => { return this.CurrentObjectName; } },
            };
            Directory = String.Empty;
            _upperTemplate = template.ToUpper();
            _needIndex = NeedToAddIndex(_upperTemplate);
            SetExtension();
        }

        public ExportFileName(string directory, string template, FileType fileType, BatchExportSettings settings)
            : this(template, fileType, settings)
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
        public string GenerateNext(dynamic sheet, dynamic selection)
        {
            Counter++;
            Logger.Info("GenerateNext: New counter: {0}", Counter);
            dynamic parent = sheet.Parent;
            CurrentWorkbookName = parent.Name;
            Bovender.ComHelpers.ReleaseComObject(parent);
            CurrentWorksheetName = sheet.Name;
            try
            {
                CurrentObjectName = selection.Name;
            }
            catch
            {
                Logger.Info("GenerateNext: Selection has no name property, using fallback 'unnamed'");
                Logger.Info("GenerateNext: Selection is a '{0}'", Microsoft.VisualBasic.Information.TypeName(selection));
                CurrentObjectName = String.Format("unnamed_{0}", Counter);
            }
            string s = _regex.Replace(Template, SubstituteVariable);
            // If no index placeholder exists in the template, add the index at the end.
            return Path.Combine(Directory, InsertIndexIfNeeded(Template, s) + _extension);
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

        private string InsertIndexIfNeeded(string template, string fileName)
        {
            if (_needIndex)
            {
                return Path.GetFileNameWithoutExtension(fileName) +
                    String.Format("({0:000})", Counter) +
                    Path.GetExtension(fileName);
            }
            else
            {
                return fileName;
            }
        }

        /// <summary>
        /// Determines whether an incremental index is required to generate unique
        /// file names.
        /// </summary>
        /// <remarks>
        /// If the layout of the objects on a worksheet is to be preserved during
        /// batch export, the file names will be unique if the template contains the
        /// worksheet name; or the workbook name and worksheet name if exporting
        /// from all workbooks, because a worksheet name may be present in several
        /// workbooks. If individual items are exported, either a name
        /// placeholder or an index placeholder must be present. An extra index
        /// is never needed if the template contains an index placeholder already.
        /// </remarks>
        /// <param name="uppercaseTemplate">File name template, converted to upper case</param>
        private bool NeedToAddIndex(string uppercaseTemplate)
        {
            if (_upperTemplate.Contains("{" + Strings.Index.ToUpper() + "}")) return false;

            bool need;
            if ((_settings != null) && (_settings.Layout == BatchExportLayout.SheetLayout))
            {
                if (_settings.Scope == BatchExportScope.OpenWorkbooks)
                {
                    need = !(_upperTemplate.Contains("{" + Strings.Workbook.ToUpper() + "}") &&
                            _upperTemplate.Contains("{" + Strings.Worksheet.ToUpper() + "}"));
                }
                else
                {
                    need = !(_upperTemplate.Contains("{" + Strings.Worksheet.ToUpper() + "}"));
                }
            }
            else
            {
                need = !(_upperTemplate.Contains("{" + Strings.Index.ToUpper() + "}") ||
                    _upperTemplate.Contains("{" + Strings.Name.ToUpper() + "}"));
            }
            return need;
        }

        private void SetExtension()
        {
            if (String.IsNullOrWhiteSpace(Template) || !_upperTemplate.EndsWith(FileType.ToFileNameExtension().ToUpper()))
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
        protected string CurrentObjectName { get; set; }
        protected FileType FileType { get; private set; }

        #endregion

        #region Private fields

        Dictionary<string, Func<string>> _placeholderReplacements;
        string _extension;
        string _upperTemplate;
        bool _needIndex;
        BatchExportSettings _settings;
        private static readonly Regex _regex = new Regex(@"{[^}]+}");

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

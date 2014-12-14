using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Export
{
    /// <summary>
    /// Generates file names to use with for graphics export.
    /// </summary>
    public class ExportFileName
    {
        #region Public properties

        public int Counter { get; protected set; }

        #endregion

        #region Constructor

        public ExportFileName(string template)
        {
            Template = template;
            Counter = 0;
            _placeholderReplacements = new Dictionary<string, Func<string>>()
            {
                { Strings.Workbook.ToUpper(), () => { return this.CurrentWorkbookName; } },
                { Strings.Worksheet.ToUpper(), () => { return this.CurrentWorksheetName; } },
                { Strings.Index.ToUpper(), () => { return String.Format("{0:000}", Counter); } },
            };
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Generates the next file name to use; this will increase the
        /// internal counter.
        /// </summary>
        /// <returns></returns>
        public string GenerateNext(Workbook workbook, dynamic worksheet)
        {
            CurrentWorkbookName = workbook.Name;
            CurrentWorksheetName = worksheet.Name;
            Counter++;
            Regex r = new Regex(@"{[^}]+}");
            string s = r.Replace(Template, SubstituteVariable);
            // If no index placeholder exists in the template, add the index at the end.
            return InsertIndexIfMissing(Template, s);
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

        #endregion

        #region Protected properties

        protected string Template { get; private set; }
        protected string CurrentWorkbookName { get; set; }
        protected string CurrentWorksheetName { get; set; }

        #endregion

        #region Private fields

        Dictionary<string, Func<string>> _placeholderReplacements;

        #endregion
    }
}

/* SheetViewModel.cs
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
using System.ComponentModel;
using System.Text.RegularExpressions;
using Bovender.Extensions;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm.ViewModels;
using XLToolbox.Excel.Models;

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// A view model for Excel sheets (worksheets, charts).
    /// </summary>
    public class SheetViewModel : ViewModelBase, IDisposable
    {
        #region Static methods

        /// <summary>
        /// Determines whether or not a sheet name must be quoted in references.
        /// If the name contains certain characters, it will be quoted.
        /// See https://www.xltoolbox.net/blog/2015/05/excel-address-syntax.html
        /// </summary>
        public static bool RequiresQuote(string sheetName)
        {
            return _charsRequiringQuote.IsMatch(sheetName);
        }

        /// <summary>
        /// Determines whether or not a sheet name and workbook path must be quoted in references.
        /// If the name contains certain characters, it will be quoted.
        /// See https://www.xltoolbox.net/blog/2015/05/excel-address-syntax.html
        /// </summary>
        public static bool RequiresQuote(string workbookPath, string sheetName)
        {
            string fn = System.IO.Path.GetFileNameWithoutExtension(workbookPath);
            return _charsRequiringQuote.IsMatch(fn) || RequiresQuote(sheetName);
        }

        #endregion

        #region Public properties

        public override string DisplayString
        {
            get
            {
                return base.DisplayString;
            }
            set
            {
                if (!IsValidName(value))
                {
                    throw new InvalidSheetNameException(
                        String.Format("The string '{0}' is not a valid sheet name",
                        value)
                    );
                };
                _sheet.Name = value;
                base.DisplayString = value;
            }
        }

        public dynamic Sheet
        {
            get
            {
                return _sheet;
            }
            set
            {
                _sheet = value;
                this.Worksheet = value as Worksheet;
                this.Chart = value as Chart;
                this.IsChart = this.Chart != null;
                if (this.Worksheet == null && this.Chart == null)
                {
                    string s = value == null ? "null value" : value.GetType().ToString();
                    throw new ArgumentException("Requires Worksheet or Chart, but not "
                        + s);
                }
                OnPropertyChanged("Sheet");
                OnPropertyChanged("MaxRows");
                OnPropertyChanged("MaxCols");
                // Set the base class' DisplayString property to prevent
                // renaming the worksheet that is triggered by writing this
                // class' DisplayString property.
                base.DisplayString = _sheet.Name;
            }
        }

        /// <summary>
        /// Returns the Sheet as a Worksheet, or Null if the Sheet is a Chart.
        /// </summary>
        public Worksheet Worksheet { get; private set; }

        /// <summary>
        /// Returns the Sheet as a Chart, or Null if the Sheet is a Worksheet.
        /// </summary>
        public Chart Chart { get; private set; }

        /// <summary>
        /// Indicates whether the Sheet model is a worksheet or a chart.
        /// This property is set by the constructor and provides quicker
        /// repeat access to the information than "myobject [AI]s Chart"
        /// statements.
        /// </summary>
        public bool IsChart { get; private set; }

        /// <summary>
        /// Returns the name of the sheet suitable for referencing.
        /// If the name contains certain characters, it will be quoted.
        /// See https://www.xltoolbox.net/blog/2015/05/excel-address-syntax.html
        /// </summary>
        public string RefName
        {
            get
            {
                if (RequiresQuote(_sheet.Name))
                {
                    return String.Format("'{0}'", _sheet.Name);
                }
                else
                {
                    return _sheet.Name;
                }
            }
        }

        public string RefNameWithWorkbook
        {
            get
            {
                string result;
                Workbook parent = _sheet.Parent;
                string path = _sheet.Parent.FullName;
                if (RequiresQuote(path, _sheet.Name))
                {
                    result = String.Format("'[{0}]{1}'", parent.Name, _sheet.Name);
                }
                else
                {
                    result = String.Format("[{0}]{1}", parent.Name, _sheet.Name);
                }
                parent.ReleaseComObject();
                return result;
            }
        }

        public string RefNameWithWorkbookAndPath
        {
            get
            {
                string result;
                Workbook parent = _sheet.Parent;
                string path = _sheet.Parent.Path;
                if (RequiresQuote(path, _sheet.Name))
                {
                    result = System.IO.Path.Combine(
                        String.Format("'{0}", path),
                        String.Format("[{0}]{1}'", parent.Name, _sheet.Name));
                }
                else
                {
                    result = System.IO.Path.Combine(
                        path,
                        String.Format("[{0}]{1}", parent.Name, _sheet.Name)
                        );
                }
                parent.ReleaseComObject();
                return result;
            }
        }

        public long MaxRows
        {
            get
            {
                long result;
                Workbook workbook = _sheet.Parent;
                if (workbook.Excel8CompatibilityMode)
                {
                    result = 2 ^ 16; // 65,536
                }
                else
                {
                    result = 2 ^ 20; // 1,048,576
                }
                workbook.ReleaseComObject();
                return result;
            }
        }

        public long MaxCols
        {
            get
            {
                long result;
                Workbook workbook = _sheet.Parent;
                if (workbook.Excel8CompatibilityMode)
                {
                    result = 2 ^ 8; // 256, "IV"
                }
                else
                {
                    result = 2 ^ 14; // 16,384, "XFD"
                }
                workbook.ReleaseComObject();
                return result;
            }
        }

        public string MaxColNotation
        {
            get
            {
                ColumnHelper c = new ColumnHelper();
                c.Number = MaxCols;
                return c.ToString();
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Counts the charts and shapes that are embedded in the worksheet,
        /// or counts the chart sheet as 1.
        /// </summary>
        /// <returns>Number of embedded charts and shapes, or 1 if the sheet
        /// is a chart.</returns>
        public int CountShapes()
        {
            if (!IsChart)
            {
                // The Shapes collection holds the chart objects as well.
                return Worksheet.Shapes.Count;
            }
            else
            {
                return 1;
            }
        }

        public int CountCharts()
        {
            if (!IsChart)
            {
                return ((Worksheet)Sheet).ChartObjects().Count;
            }
            else
            {
                return 1;
            }
        }

        /// <summary>
        /// Selects all shapes (chart and other graphic objects)
        /// on the sheet, or the chart if the sheet is a chart sheet.
        /// </summary>
        /// <returns>True if there are any charts and shapes on the
        /// sheet or if the sheet is a chart sheet; false if not.</returns>
        public bool SelectShapes()
        {
            if (!IsChart && this.Worksheet.Shapes.Count > 0)
            {
                // The Shapes collection holds the chart objects as well.
                // Select the first shape, replacing the previous selection.
                Worksheet.Shapes.Item(1).Select(true);
                foreach (Shape shape in this.Worksheet.Shapes)
                {
                    shape.Select(false);
                }
                return true;
            }
            else 
            {
                // Select the chart sheet, if any.
                return SelectCharts();
            }
        }

        /// <summary>
        /// Selects all charts on the sheet.
        /// </summary>
        /// <returns>True if the sheet is a chart or contains embedded
        /// charts; false if not.</returns>
        public bool SelectCharts()
        {
            if (IsChart)
            {
                // Cast to _Chart to prevent compile-time warning
                // about ambiguity of method and event name.
                ((_Chart)Chart).Select(true);
                return true;
            }
            else
            {
                if (this.Worksheet.ChartObjects().Count > 0)
                {
                    // Select first chart object to replace current selection.
                    // Remember that Excel collections are 1 based!
                    this.Worksheet.ChartObjects(1).Select(true);
                    foreach (ChartObject co in this.Worksheet.ChartObjects())
                    {
                        co.Select(false);
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        #endregion

        #region Constructors

        public SheetViewModel() {}

        public SheetViewModel(object sheet)
            : this()
        {
            this.Sheet = sheet;
        }

        #endregion

        #region Disposing

        ~SheetViewModel()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
                try
                {
                    ObjectExtensions.ReleaseDynamicComObject(_sheet);
                }
                catch { }
            }
        }

        #endregion

        #region Static Methods

        /// <summary>
        /// Tests whether a string represents a valid Excel sheet name.
        /// </summary>
        /// <remarks>Excel sheet names must be 1 to 31 characters long and must
        /// not contain the characters ":/\[]*?".</remarks>
        /// <param name="name">String to test.</param>
        /// <returns>True if <paramref name="name"/> can be used as a sheet name,
        /// false if not.</returns>
        public static bool IsValidName(string name)
        {
            if (!String.IsNullOrEmpty(name))
            {
                Regex r = new Regex(@"^[^:/\\*?[\]]{1,31}$");
                return r.IsMatch(name);
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _sheet;
        }

        #endregion

        #region Private fields

        private dynamic _sheet;
        private bool _disposed;
        private static readonly Regex _charsRequiringQuote = new Regex(@"\W");

        #endregion
    }
}

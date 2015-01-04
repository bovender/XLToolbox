/* SheetViewModel.cs
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
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// A view model for Excel sheets (worksheets, charts).
    /// </summary>
    public class SheetViewModel : ViewModelBase
    {
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

        #endregion
    }
}

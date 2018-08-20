/* SelectionChangedEventArgs.cs
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
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// Event args for the <see cref="SelectionChanged"/> event
    /// of the <see cref="SelectionViewModel"/> class.
    /// </summary>
    public class SelectionChangedEventArgs : EventArgs
    {
        #region Public properties

        public Worksheet ActiveWorksheet { get; set; }

        public Chart ActiveChart { get; set; }

        public Workbook ActiveWorkbook { get; set; }

        public dynamic Selection { get; set; }

        #endregion

        #region Constructors

        public SelectionChangedEventArgs() : base() { }

        public SelectionChangedEventArgs(Application excelApplication)
        {
            if (excelApplication != null)
            {
                Selection = excelApplication.Selection;
                ActiveWorkbook = excelApplication.ActiveWorkbook;
                ActiveWorksheet = excelApplication.ActiveSheet as Worksheet;
                ActiveChart = excelApplication.ActiveSheet as Chart;
            }
        }

        #endregion
    }
}

/* Address.cs
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
using System.Text.RegularExpressions;
using Xl = Microsoft.Office.Interop.Excel;
using XLToolbox.Excel.Instance;

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// Represents an address of a range (cells, columns, rows).
    /// </summary>
    public class Address : Bovender.Mvvm.ViewModels.ViewModelBase
    {
        #region Public properties

        public string Value
        {
            get
            {
                Xl.Workbook workbook = ExcelInstance.Application.ActiveWorkbook;
                if (workbook == null || workbook.Name != _workbook)
                {
                    return FullyQualified;
                }
                else
                {
                    Xl.Worksheet worksheet = workbook.ActiveSheet as Xl.Worksheet;
                    if (worksheet == null || worksheet.Name != _worksheet)
                    {
                        return QualifiedWithWorksheet;
                    }
                    else
                    {
                        return _cells;
                    }
                }
            }
            set
            {
                Regex r = new Regex(@"(.+)");
            }
        }

        /// <summary>
        /// Gets the address formatted with the qualifying worksheet.
        /// </summary>
        public string QualifiedWithWorksheet
        {
            get
            {
                return String.Format("[%s]%s", _worksheet, _cells);
            }
        }

        /// <summary>
        /// Gets the fully qualified address (with workbook and worksheet).
        /// </summary>
        public string FullyQualified
        {
            get
            {
                return String.Format("%s![%s]%s", _workbook, _worksheet, _cells);
            }
        }

        #endregion

        #region Constructors

        public Address() : base() { }

        public Address(string address)
            : this()
        {
            Value = address;
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            throw new NotImplementedException();
        }

        #endregion
 
        #region Private fields

        string _cells;
        string _worksheet;
        string _workbook;

        #endregion
   }
}

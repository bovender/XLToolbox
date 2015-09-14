/* RangeExtension.cs
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
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Excel
{
    /// <summary>
    /// Extension methods for an Excel range.
    /// </summary>
    public static class RangeExtension
    {
        /// <summary>
        /// Returns the number of cells in a range. The original Range.Count
        /// method is declared as int, which is too small for the maximum
        /// number of cells on a modern Excel worksheet
        /// (2^20 rows * 2^14 columns = 2^34 cells).
        /// </summary>
        public static long CellsCount(this Range range)
        {
            long rows = range.Rows.Count;
            long cols = range.Columns.Count;
            return rows * cols;
        }
    }
}

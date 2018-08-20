/* WorkbookExtension.cs
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
using Bovender.Extensions;

namespace XLToolbox.Excel.Extensions
{
    public static class WorkbookExtension
    {
        /// <summary>
        /// Determines if the workbook has at least one visible window.
        /// </summary>
        /// <param name="workbook">Workbook whose visibility to check.</param>
        /// <returns>True if the workbook has at least one visible window,
        /// false if not.</returns>
        public static bool IsVisible(this Workbook workbook)
        {
            Windows windows = workbook.Windows;
            Window window;
            bool visible = false;
            for (int i = 1; i <= windows.Count; i++)
            {
                window = windows[i];
                visible = window.Visible;
                Bovender.ComHelpers.ReleaseComObject(window);
                if (visible) break;
            }
            Bovender.ComHelpers.ReleaseComObject(windows);
            return visible;
        }
    }
}

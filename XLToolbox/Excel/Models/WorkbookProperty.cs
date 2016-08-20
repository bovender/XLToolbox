/* WorkbookProperty.cs
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
using System.Windows;

namespace XLToolbox.Excel.Models
{
    public class WorkbookProperty
    {
        #region Properties

        public string Name { get; set; }

        public string Value { get; set;  }

        #endregion

        #region Constructors

        public WorkbookProperty() { }

        public WorkbookProperty(string name, string value)
        {
            Name = name;
            Value = value;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Copies the value to the clipboard.
        /// </summary>
        public void Copy()
        {
            Clipboard.SetText(Value);
        }

        #endregion
    }
}

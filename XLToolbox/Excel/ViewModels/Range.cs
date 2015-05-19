/* Range.cs
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

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// View model for Excel ranges.
    /// </summary>
    /// <remarks>
    /// Supports addresses with more than 255 characters (e.g. if the range
    /// contains multiple areas).
    /// After clsMultiRange from the old XL Toolbox.
    /// </remarks>
    public class Range : Bovender.Mvvm.ViewModels.ViewModelBase
    {
        #region Public properties

        /// <summary>
        /// Gets or sets the address of the range. When getting the address,
        /// the address may be fully qualified depending on the current workbook
        /// context.
        /// </summary>
        public string Address
        {
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return Address;
        }

        #endregion

        #region Fields


        #endregion
    }
}

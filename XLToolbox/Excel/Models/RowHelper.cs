/* RowHelper.cs
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

namespace XLToolbox.Excel.Models
{
    /// <summary>
    /// Helper class to facilitate handling of row addresses.
    /// </summary>
    public class RowHelper : RangeHelperBase
    {
        #region Constructors

        public RowHelper() { }

        public RowHelper(string rowReference)
            : this()
        {
            Reference = rowReference;
        }

        public RowHelper(long rowNumber)
            : this()
        {
            Number = rowNumber;
        }

        public RowHelper(long rowNumber, bool isFixed)
            : this(rowNumber)
        {
            IsFixed = isFixed;
        }

        #endregion

        #region Implementation of RangeHelperBase

        protected override string FormatNumber(long number)
        {
            return number.ToString();
        }

        protected override long ParseNumber(string formatted)
        {
            long parsed = 0;
            if (long.TryParse(formatted, out parsed))
            {
                return parsed;
            }
            else
            {
                Logger.Fatal("ParseNumber: Failed to parse string \"{0}\"", formatted);
                throw new ArgumentException("ParseNumber: Failed to parse string");
            }
        }

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

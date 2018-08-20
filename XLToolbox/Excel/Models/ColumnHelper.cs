/* ColumnHelper.cs
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
using System.Text.RegularExpressions;

namespace XLToolbox.Excel.Models
{
    /// <summary>
    /// Helper class to facilitate handling of column addresses.
    /// </summary>
    public class ColumnHelper : RangeHelperBase
    {
        #region Constructors

        public ColumnHelper() { }

        public ColumnHelper(string columnReference)
            : this()
        {
            Reference = columnReference;
        }

        public ColumnHelper(long column)
            : this()
        {
            Number = column;
        }

        public ColumnHelper(long column, bool isFixed)
            : this()
        {
            Number = column;
            IsFixed = isFixed;
        }

        #endregion

        #region Implementation of RangeHelperBase

        protected override string FormatNumber(long number)
        {
            string s = String.Empty;
            long i = number;
            while (i > 0)
            {
                i--;
                s = Convert.ToChar(65 + i % 26) + s;
                i /= 26;
            }
            return s;
        }

        protected override long ParseNumber(string formatted)
        {
            long n = 0;
            string s = formatted.ToUpper();
            if (!_notationPattern.Value.IsMatch(s))
            {
                Logger.Fatal("ParseNumber: Invalid column notation \"{0}\"", s);
                throw new ArgumentOutOfRangeException("Invalid column notation", "formatted");
            }
            for (int i = 0; i < s.Length; i++)
            {
                n = n * 26 + Convert.ToInt32(s[i]) - 64;
            }
            return n;
        }

        #endregion

        #region Fields

        private static readonly Lazy<Regex> _notationPattern = new Lazy<Regex>(
            () => new Regex(@"^[A-Z]{1,3}$"));

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

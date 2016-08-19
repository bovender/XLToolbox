/* RangeHelperBase.cs
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
    /// Base class for the helper classes RowHelper and RangeHelper
    /// </summary>
    public abstract class RangeHelperBase : IComparable
    {
        #region Operators

        public static bool operator <(RangeHelperBase one, RangeHelperBase other)
        {
            if (one == null || other == null)
            {
                Logger.Fatal("Operator <: At least one operand is null");
                throw new InvalidOperationException("Both operands must be assigned, but at least one is null");
            }
            return one.Number < other.Number;
        }

        public static bool operator <=(RangeHelperBase one, RangeHelperBase other)
        {
            if (one == null || other == null)
            {
                Logger.Fatal("Operator <: At least one operand is null");
                throw new InvalidOperationException("Both operands must be assigned, but at least one is null");
            }
            return one.Number <= other.Number;
        }

        public static bool operator >(RangeHelperBase one, RangeHelperBase other)
        {
            if (one == null || other == null)
            {
                Logger.Fatal("Operator <: At least one operand is null");
                throw new InvalidOperationException("Both operands must be assigned, but at least one is null");
            }
            return one.Number > other.Number;
        }

        public static bool operator >=(RangeHelperBase one, RangeHelperBase other)
        {
            if (one == null || other == null)
            {
                Logger.Fatal("Operator <: At least one operand is null");
                throw new InvalidOperationException("Both operands must be assigned, but at least one is null");
            }
            return one.Number >= other.Number;
        }

        #endregion

        #region Properties

        public string Reference
        {
            get
            {
                if (_reference == null)
                {
                    string dollar = _isFixed ? "$" : "";
                    _reference = String.Format("{0}{1}", dollar, UnfixedReference);
                    Logger.Debug("Reference_get: number = {0}, fixed = {1}, reference = {2}",
                        _number, _isFixed, _reference);
                }
                return _reference;
            }
            set
            {
                if (String.IsNullOrWhiteSpace(value))
                {
                    Logger.Fatal("Reference_set: Reference must not be null or whitespace");
                    throw new ArgumentNullException("Reference must not be null or whitespace");
                }
                if (value != _reference)
                {
                    _reference = value;
                    _isFixed = _reference.StartsWith("$");
                    _number = ParseNumber(_reference.TrimStart('$'));
                    Logger.Debug("Reference_set: reference = {0}, fixed = {1}, number = {2}",
                        _reference, _isFixed, _number);
                }
            }
        }

        public string UnfixedReference
        {
            get
            {
                return FormatNumber(_number);
            }
        }

        public string FixedReference
        {
            get
            {
                return String.Format("${0}", UnfixedReference);
            }
        }

        public bool IsFixed
        {
            get
            {
                return _isFixed;
            }
            set
            {
                if (value != _isFixed)
                {
                    _isFixed = value;
                    _reference = null;
                }
            }
        }

        public long Number
        {
            get
            {
                return _number;
            }
            set
            {
                if (_number != value)
                {
                    _number = value;
                    _reference = null;
                }
            }
        }

        #endregion

        #region Public methods

        public void ToggleFixed()
        {
            IsFixed = !IsFixed;
        }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Returns a number formatted as a string.
        /// </summary>
        /// <param name="number">Number to format.</param>
        /// <returns>Formatted string.</returns>
        protected abstract string FormatNumber(long number);

        /// <summary>
        /// Parses a formatted reference string.
        /// </summary>
        /// <remarks>
        /// This method is called by the Reference setter which will strip a
        /// dollar sign (if present) from the formatted string first.
        /// </remarks>
        /// <param name="formatted">String containing a formatted number.</param>
        /// <returns>Number read from the string.</returns>
        protected abstract long ParseNumber(string formatted);

        #endregion

        #region Overrides

        public override string ToString()
        {
            return Reference;
        }

        public int CompareTo(object obj)
        {
            if (this.GetType() != obj.GetType())
            {
                throw new InvalidOperationException("Cannot compare different types");
            }
            RangeHelperBase other = obj as RangeHelperBase;
            if (this.Number < other.Number)
            {
                return -1;
            }
            else if (this.Number > other.Number)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        #endregion

        #region Private fields

        private string _reference;
        private bool _isFixed;
        private long _number;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

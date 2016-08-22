/* RangeViewModel.cs
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
using System.IO;
using Bovender.Extensions;
using System.Text.RegularExpressions;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm.ViewModels;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Excel.Models
{
    /// <summary>
    /// View model for Excel ranges.
    /// </summary>
    public class Reference : IDisposable
    {
        #region Properties

        public string WorkbookPath
        {
            get
            {
                if (_range != null)
                {
                    return _workbook.FullName;
                }
                else
                {
                    return _workbookPath;
                }
            }
            set
            {
                Reset();
                _workbookPath = value;
            }
        }

        public string SheetName
        {
            get
            {
                if (_range != null)
                {
                    return _worksheet.Name;
                }
                else
                {
                    return _worksheetName;
                }
            }
            set
            {
                Reset();
                _worksheetName = value;
            }
        }
        
        public string Address
        {
            get
            {
                if (_address == null)
                {
                    _address = Range.Address;
                    ParseAddress();
                }
                return _address;
            }
            set
            {
                Reset();
                _address = value;
                ParseAddress();
                NormalizeAddress();
            }
        }

        public string FixedAddress
        {
            get
            {
                string s = Left.FixedReference + Top.FixedReference;
                if (IsRectangle)
                {
                    s += ":" + Right.FixedReference + Bottom.FixedReference;
                }
                return s;
            }
        }

        public string UnfixedAddress
        {
            get
            {
                string s = Left.UnfixedReference + Top.UnfixedReference;
                if (IsRectangle)
                {
                    s += ":" + Right.UnfixedReference + Bottom.UnfixedReference;
                }
                return s;
            }
        }

        /// <summary>
        /// Gets or sets the fully qualified Excel reference.
        /// </summary>
        public string ReferenceString
        {
            get
            {
                if (_referenceString == null)
                {
                    string path = WorkbookPath;
                    string dir = Path.GetDirectoryName(path);
                    string wbk = Path.GetFileName(path);
                    string q = SheetViewModel.RequiresQuote(path, SheetName) ? "'" : String.Empty;
                    _referenceString = String.Format(
                        "{0}{1}{2}{0}!{3}",
                        q, 
                        Path.Combine(dir, String.Format("[{0}]", wbk)),
                        SheetName,
                        Address);
                    Logger.Info("Reference_get: Built reference: {0}", _referenceString);
                }
                return _referenceString;
            }
            set
            {
                Reset();
                _referenceString = value;
                Match m = null;
                // Prevent argument null exceptions if the reference string is empty
                if (!String.IsNullOrEmpty(_referenceString))
                {
                    m = _referencePattern.Value.Match(_referenceString);
                }
                if (m != null && m.Success)
                {
                    Logger.Info("Reference_set: valid reference");
                    IsValid = true;
                    string path = m.Groups["path"].Value;
                    string workbook = m.Groups["workbook"].Value;
                    _worksheetName = m.Groups["worksheet"].Value;
                    _address = m.Groups["address"].Value;
                    ParseAddress();
                    NormalizeAddress();
                    BuildAddress();
                    Logger.Debug("Reference_set: path = \"{0}\", workbook = \"{1}\"", path, workbook);
                    _workbookPath = Path.Combine(path, workbook);
                    if (_workbookPath.StartsWith("'") && _worksheetName.EndsWith("'"))
                    {
                        Logger.Debug("Reference_set: stripping single quotes");
                        _workbookPath = _workbookPath.Substring(1, _workbookPath.Length - 1);
                        _worksheetName = _worksheetName.Substring(0, _worksheetName.Length - 1);
                    }
                }
                else
                {
                    Logger.Info("Reference_set: invalid reference");
                    IsValid = false;
                    _workbookPath = String.Empty;
                    _worksheetName = String.Empty;
                    _address = String.Empty;
                }
            }
        }

        public bool IsValid { get; private set; }

        /// <summary>
        /// Gets or sets the range.
        /// </summary>
        public Range Range
        {
            get
            {
                if (_range == null)
                {
                    _workbook = Instance.Default.LocateWorkbook(_workbookPath);
                    if (_workbook == null)
                    {
                        Logger.Fatal("Reference_get: Unable to locate workbook");
                        throw new InvalidOperationException("Workbook does not exist");
                    }
                    _worksheet = Instance.Default.LocateWorksheet(_workbook, _worksheetName);
                    if (_worksheet == null)
                    {
                        Logger.Info("Reference_get: Using active worksheet of this workbook");
                        _worksheet = _workbook.ActiveSheet;
                    }
                    try
                    {
                        _range = _worksheet.Range[Address];
                        BuildAddress();
                        _rangeWasCreated = true;
                    }
                    catch (System.Runtime.InteropServices.COMException e)
                    {
                        Logger.Fatal("Range_get: Worksheet o.k., but invalid address \"{0}\"", Address);
                        Logger.Fatal(e);
                        throw new InvalidOperationException("Address is invalid", e);
                    }
                }
                return _range;
            }
            set
            {
                Reset();
                _range = value;
                _rangeWasCreated = false;
                if (_range != null)
                {
                    _worksheet = _range.Worksheet;
                    _workbook = _worksheet.Parent;
                    IsValid = true;
                }
                else
                {
                    IsValid = false;
                }
            }
        }

        /// <summary>
        /// Gets the Rows object for the Range. This COM object will be released
        /// when the view model is disposed.
        /// </summary>
        public Range Rows
        {
            get
            {
                if (_rows == null && Range != null)
                {
                    _rows = Range.Rows;
                }
                return _rows;
            }
        }

        /// <summary>
        /// Gets the Columns object for the Range. This COM object will be released
        /// when the view model is disposed.
        /// </summary>
        public Range Cols
        {
            get
            {
                if (_cols == null && Range != null)
                {
                    _cols = Range.Columns;
                }
                return _cols;
            }
        }

        /// <summary>
        /// Computes the number of cells in the Range; the value is cached.
        /// </summary>
        public long CellCount
        {
            get
            {
                if (_cellCount == 0 && Range != null)
                {
                    long rows = Rows.Count;
                    long cols = Cols.Count;
                    _cellCount = rows * cols;
                    Logger.Info("CellCount: {0} rows x {1} columns = {2} cells", rows, cols, _cellCount);
                }
                return _cellCount;
            }
        }

        public RowHelper Top { get; private set; }

        public RowHelper Bottom { get; private set; }

        public ColumnHelper Left { get; private set; }

        public ColumnHelper Right { get; private set; }

        public bool IsRectangle
        {
            get
            {
                return Bottom != null;
            }
        }

        #endregion

        #region Public methods

        public void Activate()
        {
            Range r = Range;
            Logger.Info("Activate: Workbook");
            ((_Workbook)_workbook).Activate();
            Logger.Info("Activate: Worksheet");
            ((_Worksheet)_worksheet).Activate();
            Logger.Info("Activate: Select range");
            r.Select();
        }

        public void MakeFixed()
        {
            Top.IsFixed = true;
            Left.IsFixed = true;
            if (IsRectangle)
            {
                Bottom.IsFixed = true;
                Right.IsFixed = true;
            }
        }

        public void MakeUnfixed()
        {
            Top.IsFixed = false;
            Left.IsFixed = false;
            if (IsRectangle)
            {
                Bottom.IsFixed = false;
                Right.IsFixed = false;
            }
        }

        public void ToggleFixed()
        {
            Top.ToggleFixed();
            Left.ToggleFixed();
            if (IsRectangle)
            {
                Bottom.ToggleFixed();
                Right.ToggleFixed();
            }
        }

        #endregion

        #region Constructors

        public Reference() : base() { }

        public Reference(string reference)
            : this()
        {
            ReferenceString = reference;
        }

        public Reference(Range range)
        {
            Range = range;
        }

        #endregion

        #region Disposal

        ~Reference()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
                if (disposing)
                {
                    // Clean up managed resources
                }
                // Clean up unmanaged resources
                Reset();
            }
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns the fully qualified reference string for the Range,
        /// including the Workbook name and the path to the workbook if
        /// the workbook has been saved previously.
        /// </summary>
        /// <remarks>
        /// Passes the request to GetFullyQualifiedReference().
        /// </remarks>
        /// <returns>Reference string for the Range</returns>
        public override string ToString()
        {
            return ReferenceString;
        }

        #endregion

        #region Private methods

        private bool Parse(string reference)
        {
            Match m = _referencePattern.Value.Match(reference);
            if (m.Success)
            {

                return true;
            }
            else
            {
                return false;
            }
        }

        private void Reset()
        {
            if (_rangeWasCreated) _range.ReleaseComObject();
            _cols = (Range)_cols.ReleaseComObject();
            _rows = (Range)_rows.ReleaseComObject();
            _workbook = (Workbook)_workbook.ReleaseComObject();
            _worksheet = (Worksheet)_worksheet.ReleaseComObject();
            _cellCount = 0;
            _referenceString = null;
            Top = null;
            Left = null;
            Bottom = null;
            Right = null;
        }

        private void BuildAddress()
        {
            if (IsRectangle)
            {
                _address = Left.Reference + Top.Reference + ":" + Right.Reference + Bottom.Reference;
            }
            else
            {
                _address = Left.Reference + Top.Reference;
            }
        }

        private void ParseAddress()
        {
            Match m = _addressPattern.Value.Match(_address);
            if (!m.Success)
            {
                Logger.Fatal("NormalizeAddress: Invalid address \"{0}\"", _address);
                throw new InvalidOperationException("Invalid address");
            }
            Top = new RowHelper(m.Groups["top"].Value);
            Left = new ColumnHelper(m.Groups["left"].Value);
            if (_address.Contains(":"))
            {
                Bottom = new RowHelper(m.Groups["bottom"].Value);
                Right = new ColumnHelper(m.Groups["right"].Value);
            }
        }

        private void NormalizeAddress()
        {
            if (IsRectangle)
            {
                if (Top > Bottom)
                {
                    RowHelper r = Top;
                    Top = Bottom;
                    Bottom = r;
                }
                if (Left > Right)
                {
                    ColumnHelper c = Left;
                    Left = Right;
                    Right = c;
                }
            }
        }

        #endregion

        #region Private fields

        private const string ADDRESS_PATTERN =
            @"(?<address>(?<left>\$?[a-zA-Z]{1,3})(?<top>\$?[0-9]{1,7})" +
            @"(:(?<right>\$?[a-zA-Z]{1,3})(?<bottom>\$?[0-9]{1,7}))?)";
        private const string WORKSHEET_PATTERN = @"(?<worksheet>[^:?*\\/[\]]{1,31})";
        private const string WORKBOOK_PATTERN = @"((?<path>.*?)?\[(?<workbook>[^]]+)])";
        private static readonly Lazy<Regex> _addressPattern =
            new Lazy<Regex>(() => new Regex("^" + ADDRESS_PATTERN + "$"));
        private static readonly Lazy<Regex> _referencePattern =
            new Lazy<Regex>(() => new Regex(@"^((" + WORKBOOK_PATTERN + ")?" + WORKSHEET_PATTERN + @"!)?" + ADDRESS_PATTERN + "$"));
        private bool _disposed;
        private bool _rangeWasCreated;
        private string _workbookPath;
        private string _worksheetName;
        private string _address;
        private string _referenceString;
        private Workbook _workbook;
        private Worksheet _worksheet;
        private Range _range;
        private Range _rows;
        private Range _cols;
        private long _cellCount;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

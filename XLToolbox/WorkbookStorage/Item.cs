/* Item.cs
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
using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using Bovender.Extensions;

namespace XLToolbox.WorkbookStorage
{
    /// <summary>
    /// Holds a single settings item for the workbook storage.
    /// </summary>
    /// <remarks>Each row on a storage worksheet constitutes one record.
    /// The first column holds the context (i.e., the name of the worksheet
    /// associated with this setting or an empty string for a setting on the
    /// global workbook level). The second column holds the name of the key.
    /// The third column holds the value.</remarks>
    internal class Item
    {
        #region Public properties

        public string Key { get; set; }
        public string Context { get; set; }
        public object Value { get; set; }

        public bool HasValue
        {
            [DebuggerStepThrough]
            get
            {
                return Value != null;
            }
        }

        #endregion

        #region Constructors
        
        /// <summary>
        /// Creates an instance from a single worksheet row.
        /// </summary>
        /// <param name="row">Row in the storage worksheet.</param>
        public Item(Worksheet sheet, int row)
        {
            ReadFromSheet(sheet, row);
        }

        /// <summary>
        /// Constructs an object from a key, context, and value.
        /// </summary>
        /// <param name="key">Item key</param>
        /// <param name="context">Worksheet context</param>
        /// <param name="value">Item value</param>
        public Item(string key, string context, object value)
        {
            this.Key = key;
            this.Context = context;
            this.Value = value;
        }

        #endregion

        #region Internal methods
        
        /// <summary>
        /// Writes out the item information to a storage sheet.
        /// </summary>
        /// <param name="_store">Storage worksheet.</param>
        /// <param name="row">Row number to write out to.</param>
        /// <param name="serializer">Instance of an XML serializer; used to convert the value
        /// to a string.</param>
        /// <returns>True if something was written, false if not.</returns>
        internal bool WriteToSheet(Worksheet sheet, int row)
        {
            if (HasValue)
            {
                Range cells = sheet.Cells;
                cells[row, 1] = Context;
                cells[row, 2] = Key;
                cells[row, 3] = Value.ToString();
                Bovender.ComHelpers.ReleaseComObject(cells);
                return true;
            }
            else
            {
                return false;
            }
        }

        internal void ReadFromSheet(Worksheet sheet, int row)
        {
            // In order to deal with the global context, we need to
            // first fetch the cell value as an object, then convert
            // it to a string using String.Format, which accepts
            // null values.
            Range cells = sheet.Cells;
            object contextValue = cells[row, 1].Value();
            Context = String.Format("{0}", contextValue);
            Key = cells[row, 2].Value2();
            Value = cells[row, 3].Value2();
            Bovender.ComHelpers.ReleaseComObject(cells);
        }

        public int AsInt()
        {
            return (int)Value;
        }

        public string AsString()
        {
            return (string)Value;
        }

        public bool AsBool()
        {
            return (bool)Value;
        }

        public T As<T>()
        {
            return (T)Value;
        }

        #endregion
    }
}

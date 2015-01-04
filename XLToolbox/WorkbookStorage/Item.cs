/* Item.cs
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
using Microsoft.Office.Interop.Excel;
using System;

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
        public string key { get; set; }
        public string context { get; set; }
        public object value { get; set; }

        /// <summary>
        /// Creates an instance from a single worksheet row.
        /// </summary>
        /// <param name="row">Row in the storage worksheet.</param>
        public Item(Worksheet sheet, int row)
        {
            ReadFromSheet(sheet, row);
        }

        public Item(string key, string context, object value)
        {
            this.key = key;
            this.context = context;
            this.value = value;
        }

        /// <summary>
        /// Writes out the item information to a storage sheet.
        /// </summary>
        /// <param name="_store">Storage worksheet.</param>
        /// <param name="row">Row number to write out to.</param>
        /// <param name="serializer">Instance of an XML serializer; used to convert the value
        /// to a string.</param>
        internal void WriteToSheet(Worksheet sheet, int row)
        {
            sheet.Cells[row, 1] = context;
            sheet.Cells[row, 2] = key;
            sheet.Cells[row, 3] = value.ToString();
        }

        internal void ReadFromSheet(Worksheet sheet, int row)
        {
            // In order to deal with the global context, we need to
            // first fetch the cell value as an object, then convert
            // it to a string using String.Format, which accepts
            // null values.
            object contextValue = sheet.Cells[row, 1].Value();
            context = String.Format("{0}", contextValue);
            key = sheet.Cells[row, 2].Value();
            value = sheet.Cells[row, 3].Value();
        }

        public int AsInt()
        {
            return (int)value;
        }

        public string AsString()
        {
            return (string)value;
        }

        public bool AsBool()
        {
            return (bool)value;
        }

        public T As<T>()
        {
            return (T)value;
        }
    }
}

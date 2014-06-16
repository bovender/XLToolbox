using Microsoft.Office.Interop.Excel;

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
            context = sheet.Cells[row, 1];
            key = sheet.Cells[row, 2];
            value = sheet.Cells[row, 3];
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.WorkbookStorage
{
    /// <summary>
    /// A collection of WorkbookStorageItems that belong to a given context
    /// (i.e., the name of the worksheet associated with the settings items,
    /// or an empty string if the settings pertain to the global workbook scope).
    /// </summary>
    internal class ContextItems : Dictionary<string, Item>
    {
    }
}

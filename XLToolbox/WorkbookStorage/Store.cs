/* Store.cs
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
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.WorkbookStorage
{
    /// <summary>
    /// Stores stuff (strings, ints, objects) in a very hidden worksheet
    /// of a workbook, and retrieves stuff from it. The store supports
    /// 'namespaces' where a worksheet can be used as a 'context', i.e.
    /// namespace. Empty context constitutes a global namespace.
    /// Uses a cache that is automatically written to the worksheet when finalizing.
    /// </summary>
    public class Store : IDisposable
    {
        #region Public properties

        /// <summary>
        /// Gets or sets the associated workbook. If a new workbook is set,
        /// writes the values to the old workbook, then reads the values from
        /// the new workbook.
        /// </summary>
        public Workbook Workbook
        {
            get
            {
                return _workbook;
            }
            set
            {
                if (Dirty)
                {
                    WriteToWorksheet();
                }
                _workbook = value;
                if (_workbook != null)
                {
                    ReadFromWorksheet();
                }
                else
                {
                    _contexts.Clear();
                }
            }
        }

        /// <summary>
        /// Sets the context of the current storage object. This may be a
        /// worksheet name or an empty string for the global context of
        /// the workbook.
        /// </summary>
        // TODO: Deal with worksheets that are renamed or deleted while the WorkbookStorage is instantiated.
        public string Context {
            get
            {
                return _context;
            }
            set
            {
                /// If context is not an empty string, it denotes a worksheet
                /// in the associated workbook. The setter will test if the workbook
                /// does indeed contain such a worksheet, throwing an exception if not.
                if (value.Length > 0) {
                    try {
                        object o = Workbook.Sheets[value];
                        _context = value;
                    }
                    catch (System.Runtime.InteropServices.COMException e) {
                        throw new InvalidContextException(
                            String.Format("Workbook has no sheet named {0}", _context), e);
                    }
                }
                else
                {
                    _context = value;
                }
            }
        }

        #endregion

        #region Protected properties

        protected Worksheet StoreSheet {
            get {
                if (_storeSheet == null) {
                    if (Workbook == null)
                    {
                        throw new WorkbookStorageException("Cannot access storage worksheet: no workbook is associated");
                    }
                    try
                    {
                        _storeSheet = Workbook.Worksheets[STORESHEETNAME];
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        bool wasSaved = Workbook.Saved;
                        dynamic previousSheet = Workbook.ActiveSheet;
                        dynamic previousSel = Workbook.Application.Selection;

                        // If the COMException is raised, the worksheet likely does not exist
                        _storeSheet = Workbook.Worksheets.Add();

                        // xlSheetVeryHidden hides the sheet so much that it cannot be made
                        // visible from the Excel graphical user interface
                        _storeSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;

                        // Give the worksheet a special name
                        _storeSheet.Name = STORESHEETNAME;

                        previousSheet.Activate();
                        previousSel.Select();
                        Workbook.Saved = wasSaved;
                    }
                }
                return _storeSheet;
            }
        }
        protected bool Dirty { get; set; }

        #endregion

        #region Private properties

        private Dictionary<string, ContextItems> _contexts;
        private const string STORESHEETNAME = "_xltb_storage_";
        private const string STORESHEETINFO = "XL Toolbox Settings";
        private const int FIRSTROW = 2;
        private string _context;
        private Workbook _workbook;
        private Worksheet _storeSheet;
        private bool _disposed = false;

        private ContextItems Items
        {
            get
            {
                if (_context == null)
                {
                    throw new UndefinedContextException();
                }
                if (_contexts.ContainsKey(_context))
                {
                    return _contexts[_context];
                }
                else
                {
                    ContextItems c = new ContextItems();
                    _contexts.Add(_context, c);
                    return c;
                }
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Instantiates the class and associates it with the active
        /// workbook of the Excel instance that is currently associated
        /// with the add-in.
        /// </summary>
        /// <exception cref="NoExcelInstanceException">If no Excel
        /// instance is running (see <see cref="ExcelInstance"/>).</exception>
        public Store()
            : this(Instance.Default.ActiveWorkbook)
        { }

        /// <summary>
        /// Instantiates the class and associates it with a workbook.
        /// </summary>
        /// <param name="workbook">Workbook object to associate the storage with.</param>
        public Store(Workbook workbook) : base()
        {
            _context = "";
            _contexts = new Dictionary<string, ContextItems>();
            this.Workbook = workbook;
        }

        #endregion

        #region Disposing and destructing

        ~Store()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                Dispose(true);
                GC.SuppressFinalize(this);
                _disposed = true;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (Dirty)
                {
                    WriteToWorksheet();
                }
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Retrieves an integer from the storage, given a key. Throws a
        /// WorkbookStorageException if the key is not found.
        /// </summary>
        /// <param name="key">Key to look up.</param>
        /// <returns>Integer value</returns>
        public int Get(string key, int def, int min, int max)
        {
            if (HasKey(key))
            {
                int i = (int)GetDynamicValue(key);
                if (i < min)
                {
                    i = min;
                }
                else if (i > max)
                {
                    i = max;
                }
                return i;
            }
            else
            {
                return def;
            }
        }

        public string Get(string key, string def)
        {
            if (HasKey(key))
            {
                return (string)GetDynamicValue(key);
            }
            else
            {
                return def;
            }
        }

        public bool Get(string key, bool def)
        {
            if (HasKey(key))
            {
                return (bool)GetDynamicValue(key);
            }
            else
            {
                return def;
            }
        }

        public T Get<T>(string key) where T : class, new()
        {
            string xml = Get(key, String.Empty);
            if (!String.IsNullOrEmpty(xml))
            {
                StringReader sr = new StringReader(xml);
                XmlSerializer xmlSer = new XmlSerializer(typeof(T));
                try
                {
                    return xmlSer.Deserialize(sr) as T;
                }
                catch
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        public void Put(string key, int i)
        {
            PutObject(key, i);
        }

        public void Put(string key, string s)
        {
            PutObject(key, s);
        }

        public void Put(string key, bool b)
        {
            PutObject(key, b);
        }

        public void Put<T>(string key, T obj) where T: class, new()
        {
            XmlSerializer xmlSer = new XmlSerializer(typeof(T));
            StringWriter sw = new StringWriter();
            xmlSer.Serialize(sw, obj);
            Put(key, sw.ToString());
        }

        /// <summary>
        /// Sets the active sheet of the current workbook as the context.
        /// </summary>
        public void UseActiveSheet()
        {
            Context = Workbook.ActiveSheet.Name;
        }

        /// <summary>
        /// Determines if the current context has a stored item with given key.
        /// </summary>
        /// <param name="key">Key to look up.</param>
        /// <returns>True if key exists in current context, false if not.</returns>
        public bool HasKey(string key)
        {
            if (key.Length == 0)
            {
                throw new EmptyKeyException();
            }
            return Items.ContainsKey(key);
        }

        /// <summary>
        /// Writes out the values to the hidden worksheet, if values were changed.
        /// </summary>
        public void Flush()
        {
            if (Dirty)
            {
                WriteToWorksheet();
            };
        }

        #endregion

        #region Protected methods

        /// <summary>
        /// Central method to put objects into the store.
        /// </summary>
        /// <param name="key">Key to store the object under.</param>
        /// <param name="o">Object to store.</param>
        protected void PutObject(string key, object o)
        {
            if (key.Length == 0)
            {
                throw new EmptyKeyException();
            };
            if (HasKey(key))
            {
                Items.Remove(key);
            };
            Item item = new Item(key, Context, o);
            Items.Add(item.key, item);
            Dirty = true;
        }

        protected dynamic GetDynamicValue(string key)
        {
            Item item;
            if (Items.TryGetValue(key, out item))
            {
                return item.value;
            }
            else
            {
                throw new UnkownKeyException(String.Format("Context {0} has no key {1}", _context, key));
            }
        }

        /// <summary>
        /// Reads all settings from the hidden storage worksheet
        /// </summary>
        protected void ReadFromWorksheet()
        {
            _contexts.Clear();
            Range r = StoreSheet.UsedRange;

            // The first row on a storage worksheet is reserved for internal
            // use (e.g., flags).
            Item item;
            ContextItems context;
            for (int row = FIRSTROW; row <= r.Rows.Count; row++)
            {
                item = new Item(_storeSheet, row);
                if (_contexts.ContainsKey(item.context))
                {
                    context = _contexts[item.context];
                }
                else
                {
                    context = new ContextItems();
                    _contexts.Add(item.context, context);
                };
                context.Add(item.key, item);
            }
        }

        /// <summary>
        /// Writes all settings to the hidden storage worksheet.
        /// </summary>
        protected void WriteToWorksheet()
        {
            // Delete the used range first
            PrepareStoreSheet();

            // Output everything to the sheet.
            int row = FIRSTROW;
            foreach (ContextItems context in _contexts.Values)
            {
                foreach (Item item in context.Values)
                {
                    item.WriteToSheet(_storeSheet, row);
                    row++;
                }
            };
            Dirty = false;
        }


        #endregion

        #region Private methods

        private void PrepareStoreSheet()
        {
            _storeSheet.UsedRange.Clear();

            // Put an informative string into the first cell;
            // this is also required in order for GetUsedRange() to return
            // the correct range.
            _storeSheet.Cells[1, 1] = STORESHEETINFO;
        }

        #endregion
    }
}

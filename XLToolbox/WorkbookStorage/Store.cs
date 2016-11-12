/* Store.cs
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
using System.Collections.Generic;
using System.IO;
using Bovender.Extensions;
using System.Xml.Serialization;
using XLToolbox.Excel.ViewModels;
using System.Diagnostics;

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
            [DebuggerStepThrough]
            get
            {
                return _workbook;
            }
            set
            {
                if (Dirty)
                {
                    Logger.Info("Workbook_set: Store is 'dirty', writing first");
                    WriteToWorksheet();
                }
                _workbook = value;
                if (_workbook != null)
                {
                    Logger.Info("Workbook_set: New workbook, reading from store sheet");
                    ReadFromWorksheet();
                }
                else
                {
                    Logger.Info("Workbook_set: No workbook, clearing contexts");
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
                if (!String.IsNullOrWhiteSpace(value)) {
                    try
                    {
                        object o = Workbook.Sheets[value];
                        Logger.Info("Context_set: Setting new worksheet context");
                        _context = value;
                    }
                    catch (System.Runtime.InteropServices.COMException e)
                    {
                        Logger.Fatal("Context_set: Invalid context '{0}' because no such sheet exists", _context);
                        Logger.Fatal(e);
                        throw new InvalidContextException(
                            String.Format("Workbook has no sheet named {0}", _context), e);
                    }
                }
                else
                {
                    Logger.Info("Context_set: Using global context");
                    _context = String.Empty;
                }
            }
        }

        #endregion

        #region Protected properties

        protected Worksheet StoreSheet
        {
            get
            {
                if (_storeSheet == null)
                {
                    if (Workbook == null)
                    {
                        Logger.Fatal("StoreSheet_Get: Workbook is null!");
                        throw new WorkbookStorageException("Cannot access storage worksheet: no workbook is associated");
                    }
                    Sheets sheets = Workbook.Worksheets;
                    try
                    {
                        _storeSheet = sheets[STORESHEETNAME];
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        CreateStoreWorksheet();
                    }
                    Bovender.ComHelpers.ReleaseComObject(sheets);
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
        {
            Logger.Info("Constructing new Store with active workbook");
        }

        /// <summary>
        /// Instantiates the class and associates it with a workbook.
        /// </summary>
        /// <param name="workbook">Workbook object to associate the storage with.</param>
        public Store(Workbook workbook) : base()
        {
            Logger.Info("Constructing new Store");
            _context = "";
            _contexts = new Dictionary<string, ContextItems>();
            this.Workbook = workbook;
            if (workbook == null)
            {
                Logger.Warn("Store: Workbook is null!");
            }
        }

        #endregion

        #region Disposing and destructing

        ~Store()
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
                    if (Dirty)
                    {
                        WriteToWorksheet();
                    }
                }
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Retrieves an integer from the storage, given a key.
        /// </summary>
        /// <param name="key">Key to look up.</param>
        /// <returns>Integer value</returns>
        public int Get(string key, int def, int min, int max)
        {
            if (HasKey(key))
            {
                Logger.Info("Get: int: got value");
                int i = Convert.ToInt32(GetDynamicValue(key));
                if (i < min)
                {
                    Logger.Warn("Get: int: returning min");
                    i = min;
                }
                else if (i > max)
                {
                    Logger.Warn("Get: int: returning max");
                    i = max;
                }
                return i;
            }
            else
            {
                Logger.Info("Get: int: returning default");
                return def;
            }
        }

        public string Get(string key, string def)
        {
            if (HasKey(key))
            {
                Logger.Info("Get: string: got value");
                return Convert.ToString(GetDynamicValue(key));
            }
            else
            {
                Logger.Info("Get: string: returning default");
                return def;
            }
        }

        public bool Get(string key, bool def)
        {
            if (HasKey(key))
            {
                Logger.Info("Get: bool: got value");
                return Convert.ToBoolean(GetDynamicValue(key));
            }
            else
            {
                Logger.Info("Get: bool: returning default");
                return def;
            }
        }

        /// <summary>
        /// Deserializes a stored object. Returns null upon failure.
        /// </summary>
        /// <typeparam name="T">Type name of the object to be deserialized.</typeparam>
        /// <param name="key">Key that this object is stored under.</param>
        /// <returns>Deserialized object of type T, or null.</returns>
        public T Get<T>(string key) where T : class, new()
        {
            Logger.Info("Get<T>: Type is {0}", typeof(T).AssemblyQualifiedName);
            string xml = Get(key, String.Empty);
            if (!String.IsNullOrEmpty(xml))
            {
                StringReader sr = new StringReader(xml);
                XmlSerializer xmlSer = new XmlSerializer(typeof(T));
                try
                {
                    return xmlSer.Deserialize(sr) as T;
                }
                catch (Exception e)
                {
                    Logger.Warn("Get<T>: Deserialization failed!");
                    Logger.Warn(e);
                    return null;
                }
            }
            else
            {
                Logger.Warn("Get<T>: No serialization string!");
                return null;
            }
        }

        public void Put(string key, int i)
        {
            Logger.Info("Put: int");
            PutObject(key, i);
        }

        public void Put(string key, string s)
        {
            Logger.Info("Put: String");
            PutObject(key, s);
        }

        public void Put(string key, bool b)
        {
            Logger.Info("Put: bool");
            PutObject(key, b);
        }

        public void Put<T>(string key, T obj) where T: class, new()
        {
            Logger.Info("Put<T>: Type is {0}", typeof(T).AssemblyQualifiedName);
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
            dynamic activeSheet = Workbook.ActiveSheet;
            Context = activeSheet.Name;
            Bovender.ComHelpers.ReleaseComObject(activeSheet);
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
                Logger.Fatal("PutObject: Empty key!");
                throw new EmptyKeyException();
            };
            if (HasKey(key))
            {
                Logger.Info("PutObject: Removing existing key '{0}'", key);
                Items.Remove(key);
            };
            Logger.Info("PutObject: Adding key '{0}'", key);
            Item item = new Item(key, Context, o);
            Items.Add(item.Key, item);
            Dirty = true;
        }

        protected dynamic GetDynamicValue(string key)
        {
            Item item;
            if (Items.TryGetValue(key, out item))
            {
                Logger.Info("GetDynamicValue: Got value for key '{0}'", key);
                return item.Value;
            }
            else
            {
                Logger.Fatal("GetDynamicValue: Unknown key '{0}' in context '{1}'", key, _context);
                throw new UnkownKeyException(String.Format("Context {0} has no key {1}", _context, key));
            }
        }

        /// <summary>
        /// Reads all settings from the hidden storage worksheet
        /// </summary>
        protected void ReadFromWorksheet()
        {
            if (Workbook == null)
            {
                Logger.Warn("ReadFromWorksheet: No workbook; exiting...");
            }

            _contexts.Clear();
            Range r = StoreSheet.UsedRange;

            // The first row on a storage worksheet is reserved for internal
            // use (e.g., flags).
            Item item;
            ContextItems context;
            Logger.Info("ReadFromWorksheet: Reading rows...");
            for (int row = FIRSTROW; row <= r.Rows.Count; row++)
            {
                item = new Item(_storeSheet, row);
                Logger.Info("ReadFromWorksheet: ... {0}::{1}", item.Context, item.Key);
                if (_contexts.ContainsKey(item.Context))
                {
                    context = _contexts[item.Context];
                }
                else
                {
                    context = new ContextItems();
                    _contexts.Add(item.Context, context);
                };
                context.Add(item.Key, item);
            }
        }

        /// <summary>
        /// Writes all settings to the hidden storage worksheet.
        /// </summary>
        protected void WriteToWorksheet()
        {
            if (Workbook == null) return;

            // Delete the used range first
            PrepareStoreSheet();

            // Output everything to the sheet.
            int row = FIRSTROW;
            foreach (ContextItems context in _contexts.Values)
            {
                if (context.Values != null)
                {
                    foreach (Item item in context.Values)
                    {
                        if (item.WriteToSheet(_storeSheet, row)) row++;
                    }
                }
            };
            Dirty = false;
        }

        /// <summary>
        /// Creates a hidden storage worksheet
        /// </summary>
        private void CreateStoreWorksheet()
        {
            if (Workbook == null)
            {
                Logger.Warn("CreateStoreWorksheet: No workbook; exiting...");
            }

            bool wasSaved = Workbook.Saved;
            dynamic previousSheet = Workbook.ActiveSheet;
            dynamic previousSel = Workbook.Application.Selection;
            Sheets sheets = Workbook.Worksheets;

            Logger.Info("CreateStoreWorksheet: Adding new sheet");
            // If the COMException is raised, the worksheet likely does not exist
            _storeSheet = sheets.Add();

            // xlSheetVeryHidden hides the sheet so much that it cannot be made
            // visible from the Excel graphical user interface
            _storeSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;

            // Give the worksheet a special name
            _storeSheet.Name = STORESHEETNAME;

            if (previousSheet != null)
            {
                Logger.Info("CreateStoreWorksheet: Activating previously active sheet");
                previousSheet.Activate();
                Bovender.ComHelpers.ReleaseComObject(previousSheet);
            }
            if (previousSel != null)
            {
                Logger.Info("CreateStoreWorksheet: Selecting previous selection");
                previousSel.Select();
                Bovender.ComHelpers.ReleaseComObject(previousSel);
            }
            Workbook.Saved = wasSaved;
            Bovender.ComHelpers.ReleaseComObject(sheets);
        }

        #endregion

        #region Private methods

        private void PrepareStoreSheet()
        {
            if (Workbook == null)
            {
                Logger.Warn("PrepareStoreSheet: No workbook; exiting...");
            }

            Logger.Info("PrepareStoreSheet: Preparing sheet");
            Range usedRange = _storeSheet.UsedRange;
            if (usedRange != null)
            {
                Logger.Info("PrepareStoreSheet: Clearing used range");
                usedRange.Clear();
                Bovender.ComHelpers.ReleaseComObject(usedRange);
            }

            // Put an informative string into the first cell;
            // this is also required in order for GetUsedRange() to return
            // the correct range.
            Range cells = _storeSheet.Cells;
            if (cells != null)
            {
                Logger.Info("PrepareStoreSheet: Writing informative header");
                cells[1, 1] = STORESHEETINFO;
                Bovender.ComHelpers.ReleaseComObject(cells);
            }
        }

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

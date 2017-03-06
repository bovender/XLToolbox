/* Instance.cs
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
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Bovender.Extensions;
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;
using XLToolbox.Excel.Extensions;

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// Provide access to an instance of Excel that the
    /// components are to work with.
    /// </summary>
    /// <remarks>
    /// <para>This class uses static fields to make sure only one
    /// instance of Excel is invoked. An internal counter records
    /// the number of class instances that are currently in use;
    /// when the last instance of this class is disposed of, the
    /// Excel instance will be closed.</para>
    /// <para>Note that this class will only reference one single
    /// Excel instance, regardless whether this was started using
    /// a static method or by instantiating the class. Thus, there
    /// is no instance property to access the Exce instance, just
    /// the static property. Instantiating this class mainly serves
    /// the purpose of being able to automatically close Excel when
    /// the work is done by using Using() structures.</para>
    /// </remarks>
    public class Instance : ViewModelBase, IDisposable
    {
        #region Singleton factory

        public static Instance Default
        {
            get { return _lazy.Value; }
            set { _lazy = new Lazy<Instance>(() => value); }
        }

        #endregion

        #region Events

        public event EventHandler<InstanceShutdownEventArgs> ShuttingDown;

        #endregion

        #region Commands

        public DelegatingCommand QuitInteractivelyCommand
        {
            get
            {
                if (_quitInteractivelyCommand == null)
                {
                    _quitInteractivelyCommand = new DelegatingCommand(
                        (param) => DoQuitInteractively());
                }
                return _quitInteractivelyCommand;
            }
        }

        public DelegatingCommand QuitSavingChangesCommand
        {
            get
            {
                if (_quitSavingChangesCommand == null)
                {
                    _quitSavingChangesCommand = new DelegatingCommand(
                        (param) => DoQuitSavingChanges(),
                        (param) => CanQuitSavingChanges());
                }
                return _quitSavingChangesCommand;
            }
        }

        public DelegatingCommand QuitDiscardingChangesCommand
        {
            get
            {
                if (_quitDiscardingChangesCommand == null)
                {
                    _quitDiscardingChangesCommand = new DelegatingCommand(
                        (param) => DoQuitDiscardingChanges(),
                        (parma) => CanQuitDiscardingChanges());
                }
                return _quitDiscardingChangesCommand;
            }
        }

        #endregion

        #region MVVM messages

        public Message<MessageContent> ConfirmQuitSavingChangesMessage
        {
            get
            {
                if (_confirmQuitSavingChangesMessage == null)
                {
                    _confirmQuitSavingChangesMessage = new Message<MessageContent>();
                }
                return _confirmQuitSavingChangesMessage;
            }
        }

        public Message<MessageContent> ConfirmQuitDiscardingChangesMessage
        {
            get
            {
                if (_confirmQuitDiscardingChangesMessage == null)
                {
                    _confirmQuitDiscardingChangesMessage = new Message<MessageContent>();
                }
                return _confirmQuitDiscardingChangesMessage;
            }
        }

        #endregion

        #region Public properties

        public Application Application
        {
            [DebuggerStepThrough]
            get
            {
                if (_application == null)
                {
                    Logger.Warn("Application_get: Returning null!");
                }
                return _application;
            }
        }

        /// <summary>
        /// Gets the Application's Workbooks collection. The underlying
        /// COM object is automatically released when the Instance
        /// object is disposed.
        /// </summary>
        public Workbooks Workbooks
        {
            get
            {
                if (_workbooks == null)
                {
                    _workbooks = Application.Workbooks;
                }
                return _workbooks;
            }
        }

        public Workbook ActiveWorkbook
        {
            get
            {
                return Application.ActiveWorkbook;
            }
        }

        /// <summary>
        /// Returns the active path. This is either the path
        /// of the active workbook, or the current working
        /// directory.
        /// </summary>
        /// <remarks>
        /// If a workbook is opened as in Protected View and
        /// is the only open workbook, Application.ActiveWorkbook
        /// will be null. Therefore this helper property was
        /// invented.
        /// </remarks>
        public string ActivePath
        {
            get
            {
                return ActiveWorkbook == null ? System.IO.Directory.GetCurrentDirectory() : ActiveWorkbook.Path;
            }
        }

        /// <summary>
        /// Gets the major version number of the Excel instance
        /// as an integer.
        /// </summary>
        public int MajorVersion
        {
            get
            {
                if (_majorVersion == 0)
	            {
                    _majorVersion = Convert.ToInt32(
                        Application.Version.Split('.')[0],
                        CultureInfo.InvariantCulture);
	            }
                return _majorVersion;
            }
        }

        /// <summary>
        /// Gets the Excel version and build number in a human-friendly form.
        /// </summary>
        /// <remarks>
        /// See http://spreadsheetpage.com/index.php/resource/excel_version_history
        /// and http://blog.pathtosharepoint.com/2014/05/06/how-to-get-your-office-365-version-number/
        /// </remarks>
        /// <param name="excel">Excel application whose version information to
        /// to retrieve.</param>
        /// <returns>String in the form of "2003", "2010 SP1" and so on.</returns>
        public string HumanFriendlyVersion
        {
            get
            {
                string versionName = String.Empty;
                string servicePack = String.Empty;
                int build = Application.Build;
                switch (MajorVersion)
                {
                    // Very old versions are ignored (won't work with VSTO anyway)
                    case 11:
                        versionName = "2003";
                        break;
                    case 12:
                        versionName = "2007";
                        // 2007 SP information: http://support.microsoft.com/kb/928116/en-us
                        if (build >= 6611) { servicePack = " SP3"; }
                        else if (build >= 6425) { servicePack = " SP2"; }
                        else if (build >= 6241) { servicePack = " SP1"; }
                        break;
                    case 14:
                        // 2010 SP information: http://support.microsoft.com/kb/2121559/en-us
                        versionName = "2010";
                        if (build >= 7015) { servicePack = " SP2"; }
                        else if (build >= 6029) { servicePack = " SP1"; }
                        break;
                    case 15:
                        // 2013 SP information: http://support.microsoft.com/kb/2817430/en-us
                        versionName = "2013";
                        if (build >= 4569) { servicePack = " SP1"; }
                        break;
                    case 16:
                        versionName = "365";
                        break; // I believe (sparse information on the web)
                }
                return String.Format("{0}{1} ({2}.{3})",
                    versionName, servicePack, Application.Version, Application.Build);
            }
        }

        public int CountOpenWorkbooks
        {
            get
            {
                return Application == null ? 0 : Workbooks.Count;
            }
        }

        public int CountUnsavedWorkbooks
        {
            get
            {
                return CountWorkbooks(wb => !wb.Saved);
            }
        }

        public int CountSavedWorkbooks
        {
            get
            {
                return CountWorkbooks(wb => wb.Saved);
            }
        }

        /// <summary>
        /// Gets whether the current Excel instance has an SDI (Excel 2013+)
        /// or not (Excel 2007/2010).
        /// </summary>
        public bool IsSingleDocumentInterface
        {
            get
            {
                return MajorVersion >= 15;
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Creates and returns a new workbook containing exactly one worksheet.
        /// </summary>
        /// <returns>Workbook with only one worksheet.</returns>
        public Workbook CreateWorkbook()
        {
            Logger.Info("CreateWorkbook");
            // Calling the Workbooks.Add method with a XlWBATemplate constant
            // creates a workbook that contains only one sheet.
            Workbook workbook = Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            return workbook;
        }

        /// <summary>
        /// Creates a workbook containing the specified number of sheets (not less than 1).
        /// </summary>
        /// <remarks>If <paramref name="numberOfSheets"/> is less than 1, the workbook will still
        /// contain one worksheet.</remarks>
        /// <param name="numberOfSheets">Number of sheets in the new workbook.</param>
        /// <returns>Workbook containing the specified number of sheets (not less than 1).</returns>
        public Workbook CreateWorkbook(int numberOfSheets)
        {
            Logger.Info("CreateWorkbook({0})", numberOfSheets);
            Workbook wb = CreateWorkbook();
            Sheets sheets = wb.Sheets;
            for (int i = 2; i <= numberOfSheets; i++)
            {
                sheets.Add(After: sheets[sheets.Count]);
            };
            Bovender.ComHelpers.ReleaseComObject(sheets);
            return wb;
        }

        /// <summary>
        /// Locates a workbook and returns it. If the workbook is loaded already,
        /// it will not be re-opened. If the workbook cannot be located, this method
        /// returns null.
        /// </summary>
        /// <param name="name">Name (either bare name or full name with path) of the
        /// workbook being sought. If this is null or whitespace, the active workbook
        /// is returned.</param>
        /// <returns>Workbook or null</returns>
        public Workbook LocateWorkbook(string name)
        {
            Predicate<Workbook> predicate;

            if (Path.GetFileName(name) == name)
            {
                // Find workbook by name only
                predicate = new Predicate<Workbook>(wb => wb.Name == name);
            }
            else
            {
                // Find workbook by full name (including path)
                predicate = new Predicate<Workbook>(wb => wb.FullName == name);
            }

            Workbook foundWorkbook = null;
            if (!String.IsNullOrWhiteSpace(name))
            {
                Logger.Info("LocateWorkbook: Locating {0}", name);
                for (int i = 1; i <= Workbooks.Count; i++)
                {
                    Workbook workbook = Workbooks[i];
                    if (predicate(workbook))
                    {
                        Logger.Info("LocateWorkbook: Workbook already loaded");
                        foundWorkbook = workbook;
                        break;
                    }
                    Bovender.ComHelpers.ReleaseComObject(workbook);
                }
                if (foundWorkbook == null)
                {
                    Logger.Info("LocateWorkbook: Workbook not loaded, attempting to open it");
                    try
                    {
                        foundWorkbook = Workbooks.Open(name);
                    }
                    catch (Exception e)
                    {
                        Logger.Warn("LocateWorkbook: Failed to open workbook");
                        Logger.Warn(e);
                    }
                }
            }
            else
            {
                Logger.Info("LocateWorkbook: Using active workbook");
                foundWorkbook = Instance.Default.ActiveWorkbook;
            }
            if (foundWorkbook == null)
            {
                Logger.Warn("LocateWorkbook: Unable to locate workbook");
            }
            return foundWorkbook;
        }

        /// <summary>
        /// Locates a worksheet in a given workbook.
        /// </summary>
        /// <param name="workbook">Workbook in which to look for the worksheet.
        /// If this is null, the active workbook is used.</param>
        /// <param name="name">Name of the worksheet.</param>
        /// <returns>Worksheet or null if no worksheet by this name exists.</returns>
        public Worksheet LocateWorksheet(Workbook workbook, string name)
        {
            Worksheet worksheet = null;
            bool createdCOMObject = false;
            if (workbook == null)
            {
                Logger.Info("LocateWorksheet: Workbook is null, using active workbook");
                workbook = Application.ActiveWorkbook;
                createdCOMObject = true;
            }
            if (workbook == null)
            {
                Logger.Fatal("LocateWorksheet: No active workbook present, cannot proceed");
                throw new ArgumentNullException(
                    "Unable to locate worksheet because no workbook is given and there is no active workbook either",
                    "workbook");
            }
            Sheets worksheets = workbook.Worksheets;
            try
            {
                Logger.Info("LocateWorksheet: Locating \"{0}\"", name);
                worksheet = worksheets[name];
            }
            catch (Exception e)
            {
                Logger.Warn("LocateWorksheet: Unable to locate worksheet, returning null");
                Logger.Warn(e);
            }
            Bovender.ComHelpers.ReleaseComObject(worksheets);
            if (createdCOMObject) Bovender.ComHelpers.ReleaseComObject(workbook);
            return worksheet;
        }

        /// <summary>
        /// Disables screen updating. Increases an internal counter
        /// to be able to handle cascading calls to this method.
        /// </summary>
        public void DisableScreenUpdating()
        {
            if (_preventScreenUpdating == 0)
            {
                _wasScreenUpdating = Application.ScreenUpdating;
                Logger.Info("Disable screen updating");
            }
            Application.ScreenUpdating = false;
            _preventScreenUpdating++;
        }

        /// <summary>
        /// Decreases the internal screen updating counter by one;
        /// if the counter reaches 0, the application's screen updating
        /// will resume.
        /// </summary>
        public void EnableScreenUpdating()
        {
            _preventScreenUpdating--;
            if (_preventScreenUpdating <= 0)
            {
                Logger.Info("Enable screen updating");
                _preventScreenUpdating = 0;
                Application.ScreenUpdating = _wasScreenUpdating;
            }
        }

        /// <summary>
        /// Disables displaying of user alerts. Increases an internal counter
        /// to be able to handle cascading calls to this method.
        /// </summary>
        public void DisableDisplayAlerts()
        {
            if (_disableDisplayAlerts == 0)
            {
                Logger.Info("Disable displaying of alerts");
                _wasDisplayingAlerts = Application.DisplayAlerts;
            }
            Application.DisplayAlerts = false;
            _disableDisplayAlerts++;
        }

        /// <summary>
        /// Decreases the internal screen updating counter by one;
        /// if the counter reaches 0, the application's display of
        /// user alerts will be turned on again (in fact, it will
        /// be reset to its original state).
        /// </summary>
        public void EnableDisplayAlerts()
        {
            _disableDisplayAlerts--;
            if (_disableDisplayAlerts <= 0)
            {
                Logger.Info("Enable displaying of alerts");
                _disableDisplayAlerts = 0;
                Application.DisplayAlerts = _wasDisplayingAlerts;
            }
        }

        /// <summary>
        /// Debug method to reset the Excel application. The result
        /// is an application without open workbooks.
        /// </summary>
        [Conditional("DEBUG")]
        public void Reset()
        {
            DisableDisplayAlerts();
            foreach (Workbook wb in Application.Workbooks)
            {
                wb.Close();
            }
            Application.DisplayAlerts = true;
            _disableDisplayAlerts = 0;
            Application.ScreenUpdating = true;
            _preventScreenUpdating = 0;
        }

        /// <summary>
        /// Fetches a workbook if it is opened. If no workbook is found
        /// by the given name, this function returns null. Unlike the
        /// LocateWorkbook method, this method will not open a workbook
        /// that is currently not loaded.
        /// </summary>
        /// <param name="workbookName">Workbook to fetch.</param>
        /// <returns>Workbook or null.</returns>
        public Workbook FindWorkbook(string workbookName)
        {
            Workbook wb = null;
            try
            {
                Logger.Debug("FindWorkbook: Looking for {0}", workbookName);
                wb = Workbooks[workbookName];
            }
            catch (Exception e)
            {
                Logger.Debug("FindWorkbook: Caught an exception");
                Logger.Debug(e);
                Logger.Debug("FindWorkbook: Evidently \"{0}\" is not open", workbookName);
            }
            return wb;
        }

        /// <summary>
        /// Fetches an add-in if it is opened. If no add-in is found
        /// by the given name, this function returns null.
        /// </summary>
        /// <param name="addInName">Add-in to fetch.</param>
        /// <returns>Workbook or null.</returns>
        public AddIn FindAddIn(string addInName)
        {
            AddIns2 addins = Application.AddIns2;
            AddIn a = null;
            try
            {
                a = addins[addInName];
            }
            catch { }
            Bovender.ComHelpers.ReleaseComObject(addins);
            return a;
        }

        /// <summary>
        /// Returns true if a workbook is opened.
        /// </summary>
        /// <param name="addInName">Workbook name to query.</param>
        /// <returns>True if the workbook is opened.</returns>
        public bool IsWorkbookLoaded(string workbookName)
        {
            return FindWorkbook(workbookName) != null;
        }

        /// <summary>
        /// Returns true if an add-in is loaded.
        /// </summary>
        /// <param name="workbookName">Add-in name to query.</param>
        /// <returns>True if the add-in is loaded.</returns>
        public bool IsAddInLoaded(string addInName)
        {
            return FindAddIn(addInName) != null;
        }

        /// <summary>
        /// Loads an embedded resource add-in.
        /// </summary>
        /// <param name="resourceName">Addin as 'embedded resource'</param>
        /// <returns>File name of the temporary file that the resource
        /// was written to.</returns>
        internal string LoadAddinFromEmbeddedResource(string resourceName)
        {
            Stream resourceStream = typeof(Instance).Assembly
                .GetManifestResourceStream(resourceName);
            if (resourceStream == null)
            {
                Logger.Error("LoadAddinFromEmbeddedResource: Unable to read embedded resource '{0}'", resourceName);
                throw new IOException("Unable to open resource stream " + resourceName);
            }
            string addinPath;
            Workbook loadedAddin = FindWorkbook(resourceName);
            if (loadedAddin == null)
            {
                string tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                Directory.CreateDirectory(tempDir);
                addinPath = Path.Combine(tempDir, resourceName);
                Stream tempStream = File.Create(addinPath);
                resourceStream.CopyTo(tempStream);
                tempStream.Close();
                resourceStream.Close();
                try
                {
                    Logger.Info("LoadAddinFromEmbeddedResource: Loading...");
                    Workbooks.Open(addinPath);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Logger.Warn("LoadAddinFromEmbeddedResource: COM exception caught, falling back to CorruptLoad");
                    DisableDisplayAlerts();
                    try
                    {
                        Workbooks.Open(addinPath, CorruptLoad: XlCorruptLoad.xlExtractData);
                    }
                    catch (System.Runtime.InteropServices.COMException e)
                    {
                        Logger.Fatal("LoadAddinFromEmbeddedResource: COM exception occurred after calling Workbooks.Open");
                        Logger.Fatal(e);
                        throw new XLToolbox.Excel.ExcelException("Excel failed to load the legacy Toolbox add-in", e);
                    }
                    finally
                    {
                        EnableDisplayAlerts();
                    }
                }

                Logger.Info("LoadAddinFromEmbeddedResource: Loaded {0}", addinPath);
            }
            else
            {
                addinPath = loadedAddin.FullName;
                Logger.Info("LoadAddinFromEmbeddedResource: Already loaded, path is {0}", addinPath);
            }
            return addinPath;
        }

        /// <summary>
        /// Quits the current instance of Excel; no warning message will be shown.
        /// </summary>
        public void Quit()
        {
            if (_application != null)
            {
                Logger.Info("Shutdown");
                _canQuitExcel = true;
                Dispose();
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Instantiates this class without an Excel instance
        /// </summary>
        public Instance() : base() { }

        /// <summary>
        /// Creates a new instance using <paramref name="application"/> as Excel
        /// instance.
        /// </summary>
        /// <param name="application">Excel instance.</param>
        public Instance(Application application)
            : this()
        {
            if (application == null)
            {
                Logger.Warn("Instance(Application): Got null value!");
            }
            _application = application;
        }

        private Instance(bool createNewExcelInstance)
            : this()
        {
            if (createNewExcelInstance)
            {
                _canQuitExcel = true;
                _application = new Application();
                _createdInstance = true;
            }
        }

        #endregion

        #region Disposing

        ~Instance()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Logger.Info("Dispose");
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
                    OnShuttingDown();
                }
                if (_canQuitExcel && _application != null)
                {
                    _application.DisplayAlerts = false;
                    _application.Quit();
                    _workbooks = (Workbooks)Bovender.ComHelpers.ReleaseComObject(_workbooks);
                    if (_createdInstance)
                    {
                        Bovender.ComHelpers.ReleaseComObject(_application);
                    }
                    _application = null;
                }
            }
        }

        #endregion

        #region ViewModelBase implementation

        public override object RevealModelObject()
        {
            return _application;
        }

        #endregion

        #region Private methods

        private void OnShuttingDown()
        {
            EventHandler<InstanceShutdownEventArgs> h = ShuttingDown;
            if (h != null)
            {
                Logger.Info("OnShuttingDown: {0} event subscriber(s)", h.GetInvocationList().Length);
                h(this, new InstanceShutdownEventArgs(Application));
            }
            else
            {
                Logger.Info("OnShuttingDown: No-one is listening.");
            }
        }

        private void DoQuitInteractively()
        {
            Logger.Info("DoQuitInteractively");
            DoCloseView();
            CloseAllWorkbooksThenShutdown();
        }

        private void DoQuitSavingChanges()
        {
            Logger.Info("DoQuitSavingChanges");
            ConfirmQuitSavingChangesMessage.Send(
                new MessageContent(),
                (MessageContent response) =>
                {
                    if (response.Confirmed) ConfirmQuitSavingChanges();
                });
        }

        /// <summary>
        /// Called by <see cref="DoQuitSavingChanges"/> if the view responded
        /// affirmatory to the <see cref="ConfirmQuiSavingChangesMessage"/>.
        /// </summary>
        private void ConfirmQuitSavingChanges()
        {
            DoCloseView();
            Logger.Info("ConfirmQuitSavingChanges");
            Workbook w = null;
            bool canQuit = true;
            for (int i = 1; i <= Workbooks.Count; i++)
            {
                w = Workbooks[i];
                if (w.Path == String.Empty)
                {
                    Logger.Info("ConfirmQuitSavingChanges: Workbook #{0} of {1} has no path, invoking xlDialogSaveAs",
                        i, Workbooks.Count);
                    // Cast to prevent ambiguity
                    ((_Workbook)w).Activate();
                    Application.Dialogs[XlBuiltInDialog.xlDialogSaveAs].Show();
                }
                else
                {
                    Logger.Info("ConfirmQuitSavingChanges: Workbook #{0} of {1} has a path, calling Save()",
                        i, Workbooks.Count);
                    w.Save();
                }
                bool saved = w.Saved;
                Bovender.ComHelpers.ReleaseComObject(w);
                if (!saved)
                {
                    Logger.Warn("ConfirmQuitSavingChanges: Workbook #{0} of {1} not saved: not quitting",
                        i, Workbooks.Count);
                    canQuit = false;
                    break;
                }
            }
            if (canQuit)
            {
                Logger.Info("ConfirmQuitSavingChanges: Proceeding to shutdown");
                CloseAllWorkbooksThenShutdown();
            }
        }

        private bool CanQuitSavingChanges()
        {
            return CountUnsavedWorkbooks > 0;
        }

        private void DoQuitDiscardingChanges()
        {
            Logger.Info("DoQuitDiscardingChanges");
            ConfirmQuitDiscardingChangesMessage.Send(
                new MessageContent(),
                (MessageContent response) =>
                {
                    if (response.Confirmed) ConfirmQuitDiscardingChanges();
                });
        }

        /// <summary>
        /// Called by <see cref="DoQuitDiscardingChanges"/> if the view responded
        /// affirmatory to the <see cref="ConfirmQuitDiscardingChangesMessage"/>.
        /// </summary>
        private void ConfirmQuitDiscardingChanges()
        {
            Logger.Info("ConfirmQuitDiscardingChanges");
            Workbook w;
            for (int i = 1; i <= Workbooks.Count; i++)
            {
                w = Workbooks[i];
                w.Saved = true;
                if (!w.Saved)
                {
                    Logger.Warn("ConfirmQuitDiscardingChanges: Workbook #{0} of {1} still not saved!",
                        i, Workbooks.Count);
                }
                Bovender.ComHelpers.ReleaseComObject(w);
            }
            CloseAllWorkbooksThenShutdown();
        }

        private bool CanQuitDiscardingChanges()
        {
            return CountUnsavedWorkbooks > 0;
        }

        /// <summary>
        /// Closes all workbooks.
        /// </summary>
        /// <returns>True if all workbooks were closed, false if not.</returns>
        private bool CloseAllWorkbooksThenShutdown()
        {
            DoCloseView();
            // Must use a task in order to prevent hangs.
            System.Threading.Tasks.Task.Factory.StartNew((System.Action)(() =>
            {
                Logger.Info("CloseAllWorkbooksThenShutdown");
                bool allClosed = WorkWithVisibleWorkbooks(
                    wb =>
                    {
                        int oldCount = Workbooks.Count;
                        wb.Close();
                        return Workbooks.Count == oldCount - 1;
                    });
                if (allClosed)
                {
                    Logger.Info("CloseAllWorkbooksThenShutdown: Now quitting...");
                    // Call the Quit method on the Application object rather than the Quit method
                    // on the ViewModel to let others have a chance to use the ViewModel during
                    // the shutdown process.
                    _application.Quit();
                }
                else
                {
                    Logger.Info("CloseAllWorkbooksThenShutdown: At least one workbook was not closed; not shutting down.");
                }
            }));
            return true;
        }

        private int CountWorkbooks(Predicate<Workbook> test)
        {
            if (Application != null)
            {
                int n = 0;
                Workbook w;
                for (int i = 1; i <= Workbooks.Count; i++)
                {
                    w = Workbooks[i];
                    if (test(w)) n++;
                    Bovender.ComHelpers.ReleaseComObject(w);
                }
                return n;
            }
            else
            {
                return 0;
            }
        }

        private bool WorkWithVisibleWorkbooks(Predicate<Workbook> operation)
        {
            Logger.Debug("WorkWithVisibleWorkbooks: Iterating backwards...");
            Workbook w;
            bool success = true;
            int count = Workbooks.Count;
            // Work backwards because a workbook may vanish
            for (int i = count; i >= 1; i--)
            {
                Logger.Debug("WorkWithVisibleWorkbooks: Processing #{0} of {1}", i, count);
                w = Workbooks[i];
                if (w.IsVisible()) success = operation(w);
                Bovender.ComHelpers.ReleaseComObject(w);
                if (!success) break;
            }
            Logger.Debug("WorkWithVisibleWorkbooks: Success: {0}", success);
            return success;
        }

        #endregion
       
        #region Private instance fields

        private bool _disposed;
        private bool _createdInstance;
        private bool _canQuitExcel;
        private Application _application;
        private Workbooks _workbooks;
        private int _majorVersion;
        private DelegatingCommand _quitInteractivelyCommand;
        private DelegatingCommand _quitSavingChangesCommand;
        private DelegatingCommand _quitDiscardingChangesCommand;
        private Message<MessageContent> _confirmQuitSavingChangesMessage;
        private Message<MessageContent> _confirmQuitDiscardingChangesMessage;
        private bool _wasScreenUpdating;
        private bool _wasDisplayingAlerts;
        private int _preventScreenUpdating;
        private int _disableDisplayAlerts;

        #endregion

        #region Private static fields

        private static Lazy<Instance> _lazy = new Lazy<Instance>(
            () =>
            {
                Instance i = new Instance(true);
                Workbook w = i.Workbooks.Add();
                Bovender.ComHelpers.ReleaseComObject(w);
                return i;
            }
        );

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

/* WorkbookViewModel.cs
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
using System.ComponentModel;
using System.Linq;
using System.Collections.ObjectModel;
using System.Threading;
using System.Windows;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;
using Bovender.Extensions;
using XLToolbox.Excel.Models;

namespace XLToolbox.Excel.ViewModels
{
    /// <summary>
    /// View model for an Excel workbook containing a list of sheets (worksheets, charts)
    /// that can be managed (moved around, added, deleted, renamed).
    /// </summary>
    public class WorkbookViewModel : ViewModelBase
    {
        #region Public properties

        public int NumSelectedSheets { get; private set; }

        public ObservableCollection<SheetViewModel> Sheets
        {
            get
            {
                return _sheets;
            }
            protected set
            {
                _sheets = value;
                OnPropertyChanged("Sheets");
            }
        }

        public SheetViewModel ActiveSheet
        {
            get
            {
                if (_workbook != null)
                {
                    dynamic s = _workbook.ActiveSheet;
                    int i = IndexOf(s);
                    ObjectExtensions.ReleaseDynamicComObject(s);
                    return Sheets[i];
                }
                else
                {
                    return null;
                }
            }
        }

        public Message<MessageContent> ConfirmDeleteMessage
        {
            get
            {
                if (_confirmDeleteMessage == null)
                {
                    _confirmDeleteMessage = new Message<MessageContent>();
                };
                return _confirmDeleteMessage;
            }
        }

        public Message<StringMessageContent> RenameSheetMessage
        {
            get
            {
                if (_renameSheetMessage == null)
                {
                    _renameSheetMessage = new Message<StringMessageContent>();
                };
                return _renameSheetMessage;
            }
        }

        public bool AlwaysOnTop
        {
            get
            {
                return UserSettings.UserSettings.Default.WorksheetManagerAlwaysOnTop;
            }
            set
            {
                UserSettings.UserSettings.Default.WorksheetManagerAlwaysOnTop = value;
            }
        }

        /// <summary>
        /// Concatenates the sheet names in the workbook.
        /// </summary>
        public string SheetsString
        {
            get
            {
                string s = String.Empty;
                if (_workbook != null)
                {
                    Sheets sheets = null;
                    try
                    {
                        sheets = _workbook.Sheets;
                        for (int i = 1; i <= sheets.Count; i++)
                        {
                            dynamic sheet = sheets[i];
                            // Use colon as separator because it is one of the
                            // characters that are illegal in a sheet name.
                            s += sheet.Name + "::";
                            // ObjectExtensions.ReleaseDynamicComObject(sheet);
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        return _lastSheetsString;
                    }
                    finally
                    {
                        sheets.ReleaseComObject();
                    }
                }
                return s;
            }
        }

        public IList<WorkbookProperty> Properties
        {
            get
            {
                if (_properties == null && _workbook != null)
                {
                    _properties = new List<WorkbookProperty>();
                    SelectionViewModel selection = new SelectionViewModel(Instance.Default.Application);
                    if (selection.IsRange)
                    {
                        _properties.Add(new WorkbookProperty(Strings.Selection, selection.Reference.ReferenceString));
                    }
                    else
                    {
                        try
                        {
                            _properties.Add(new WorkbookProperty(Strings.Selection, selection.Selection.Name));
                        }
                        catch
                        {
                            _properties.Add(new WorkbookProperty(Strings.Selection, String.Empty));
                        }
                    }
                    _properties.Add(new WorkbookProperty(Strings.Workbook, _workbook.Name));
                    _properties.Add(new WorkbookProperty(Strings.Folder, _workbook.Path));
                    _properties.Add(new WorkbookProperty(Strings.WorkbookAndPath, _workbook.FullName));
                }
                return _properties;
            }
        }

        public int PreferredPropertyIndex
        {
            get
            {
                return UserSettings.UserSettings.Default.PreferredPropertyIndex;
            }
            set
            {
                UserSettings.UserSettings.Default.PreferredPropertyIndex = value;
                OnPropertyChanged("PreferredPropertyIndex");
            }
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Finds the index of a given sheet (worksheet or chart sheet)
        /// in the Sheets collection.
        /// </summary>
        /// <param name="sheet">Worksheet or chart sheet</param>
        /// <returns>Index of the sheet in the Sheets collection</returns>
        public int IndexOf(dynamic sheet)
        {
            string name = sheet.Name;
            SheetViewModel svm = Sheets.FirstOrDefault(s => s.DisplayString == name);
            if (svm == null)
            {
                // If the sheet was not found, rebuild the sheet list.
                // Maybe the sheet was just added, before the watch timer got active.
                BuildSheetList();
                svm = Sheets.FirstOrDefault(s => s.DisplayString == name);
            }
            if (svm != null)
            {
                Logger.Info("IndexOf: Found view model for this sheet");
                return Sheets.IndexOf(svm);
            }
            else
	        {
                Logger.Warn("IndexOf: Requested sheet not found in collection!");
                return -1;
	        }
        }

        #endregion

        #region Commands

        public DelegatingCommand MoveSheetUp
        {
            get
            {
                if (_moveSheetUp == null)
                {
                    _moveSheetUp = new DelegatingCommand(
                        parameter => { DoMoveSheetUp(); },
                        parameter => { return CanMoveSheetUp(); },
                        this).ListenOn("Workbook").ListenOn("ActiveSheet");
                }
                return _moveSheetUp;
            }
        }

        public DelegatingCommand MoveSheetsToTop
        {
            get
            {
                if (_moveSheetsToTop == null)
                {
                    _moveSheetsToTop = new DelegatingCommand(
                        parameter => { DoMoveSheetsToTop(); },
                        parameter => { return CanMoveSheetsToTop(); },
                        this).ListenOn("Workbook").ListenOn("ActiveSheet");
                }
                return _moveSheetsToTop;
            }
        }

        public DelegatingCommand MoveSheetDown
        {
            get
            {
                if (_moveSheetDown == null)
                {
                    _moveSheetDown = new DelegatingCommand(
                        parameter => { DoMoveSheetDown(); },
                        parameter => { return CanMoveSheetDown(); },
                        this).ListenOn("Workbook").ListenOn("ActiveSheet");
                }
                return _moveSheetDown;
            }
        }

        public DelegatingCommand MoveSheetsToBottom
        {
            get
            {
                if (_moveSheetsToBottom == null)
                {
                    _moveSheetsToBottom = new DelegatingCommand(
                        parameter => { DoMoveSheetsToBottom(); },
                        parameter => { return CanMoveSheetsToBottom(); },
                        this).ListenOn("Workbook").ListenOn("ActiveSheet");
                }
                return _moveSheetsToBottom;
            }
        }

        public DelegatingCommand DeleteSheets
        {
            get
            {
                if (_deleteSheets == null)
                {
                    _deleteSheets = new DelegatingCommand(
                        parameter => { DoDeleteSheets(); },
                        parameter => { return CanDeleteSheets(); },
                        this).ListenOn("Workbook").ListenOn("ActiveSheet");
                }
                return _deleteSheets;
            }
        }

        public DelegatingCommand RenameSheet
        {
            get
            {
                if (_renameSheet == null)
                {
                    _renameSheet = new DelegatingCommand(
                        parameter => { DoRenameSheet(); },
                        parameter => { return CanRenameSheet(); },
                        this).ListenOn("Workbook").ListenOn("ActiveSheet");
                }
                return _renameSheet;
            }
        }

        /// <summary>
        /// Monitors the current workbook by periodically checking the list
        /// of sheet names.
        /// </summary>
        /// <remarks>
        /// This is a workaround for the absence of a "Workbook changed", "Sheet
        /// order changed" or similar event in Excel.
        /// </remarks>
        public DelegatingCommand MonitorWorkbook
        {
            get
            {
                if (_monitorWorkbook == null)
                {
                    _monitorWorkbook = new DelegatingCommand(
                        parameter => { DoMonitorWorkbook(); },
                        parameter => { return CanMonitorWorkbook(); },
                        this).ListenOn("Workbook").ListenOn("ActiveSheet");
                }
                return _monitorWorkbook;
            }
        }

        /// <summary>
        /// Stops monitoring the current workbook for changes in its sheet
        /// list.
        /// </summary>
        public DelegatingCommand UnmonitorWorkbook
        {
            get
            {
                if (_unmonitorWorkbook == null)
                {
                    _unmonitorWorkbook = new DelegatingCommand(
                        parameter => { DoUnmonitorWorkbook(); },
                        parameter => { return CanUnmonitorWorkbook(); },
                        this).ListenOn("Workbook").ListenOn("ActiveSheet");
                }
                return _unmonitorWorkbook;
            }
        }

        public DelegatingCommand CopyPropertyCommand
        {
            get
            {
                if (_copyPropertyCommand == null)
                {
                    _copyPropertyCommand = new DelegatingCommand(CopyProperty, CanCopyProperty);
                }
                return _copyPropertyCommand;
            }
        }

        #endregion

        #region Constructors

        public WorkbookViewModel()
        {
            if (!XLToolbox.Excel.ViewModels.Instance.Default.IsSingleDocumentInterface)
            {
                // Change the workbook model only if this is not an SDI application
                Excel.ViewModels.Instance.Default.Application.WorkbookActivate += Application_WorkbookActivate;
                Excel.ViewModels.Instance.Default.Application.WorkbookDeactivate += Application_WorkbookDeactivate;
            }
            Instance.Default.ShuttingDown += (sender, args) =>
            {
                DoUnmonitorWorkbook();
            };
        }

        public WorkbookViewModel(Workbook workbook)
            : this()
        {
            this.Workbook = workbook;
        }

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _workbook;
        }

        #endregion

        #region Protected methods

        protected void BuildSheetList()
        {
            NumSelectedSheets = 0;
            if (Workbook != null)
            {
                _lastSheetsString = SheetsString;
                ObservableCollection<SheetViewModel> sheetViewModels = new ObservableCollection<SheetViewModel>();
                SheetViewModel svm;
                Sheets excelSheets = Workbook.Sheets;
                for (int i = 1; i <= excelSheets.Count; i++)
                {
                    dynamic sheet = excelSheets[i];
                    // Need to cast because directly comparing the Visible property with
                    // XlSheetVisibility.xlSheetVisible caused exceptions.
                    if ((XlSheetVisibility)sheet.Visible == XlSheetVisibility.xlSheetVisible)
                    {
                        svm = new SheetViewModel(sheet);
                        svm.PropertyChanged += svm_PropertyChanged;
                        sheetViewModels.Add(svm);
                    }
                    ObjectExtensions.ReleaseDynamicComObject(sheet);
                };
                excelSheets.ReleaseComObject();
                Sheets = sheetViewModels;
                dynamic activeSheet = Workbook.ActiveSheet;
                if (activeSheet != null)
                {
                    Logger.Info("BuildSheetList: Selecting active sheet in list");
                    SheetActivated(activeSheet);
                    ObjectExtensions.ReleaseDynamicComObject(activeSheet);
                }
                else
                {
                    Logger.Info("BuildSheetList: Cannot select active sheet in list; ActiveSheet is null.");
                }
            }
            else
            {
                Sheets = null;
            }
        }

        #endregion

        #region Event handlers

        private void svm_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsSelected")
            {
                SheetViewModel svm = sender as SheetViewModel;
                if (svm.IsSelected)
                {
                    NumSelectedSheets++;
                    _lastSelectedSheet = svm;
                    if (_lockEvents <= 0)
                    {
                        svm.Sheet.Activate();
                    }
                }
                else
                {
                    NumSelectedSheets--;
                }
            }
        }

        void Application_WorkbookActivate(Workbook Wb)
        {
            Workbook = Wb;
        }

        void Application_WorkbookDeactivate(Workbook Wb)
        {
            Workbook = null;
        }

        #endregion

        #region Private methods

        private void DoMoveSheetUp()
        {
            _lockEvents += 1;
            // When iterating over the worksheet view models in the Sheets collection
            // as well as over the sheets collection of the workbook, keep in mind
            // that Excel workbook collections are 1-based.
            for (int i = 1; i < Sheets.Count; i++)
            {
                Sheets sheets = Workbook.Sheets;
                if (Sheets[i].IsSelected)
                {
                    var moving = sheets[i + 1];
                    var other = sheets[i];
                    moving.Move(before: other);
                    moving.ReleaseComObject();
                    other.ReleaseComObject();
                    Sheets.Move(i, i - 1);
                }
                sheets.ReleaseComObject();
            }
            _lockEvents -= 1;
        }

        private void DoMoveSheetsToTop()
        {
            _lockEvents += 1;
            int currentTop = 0;
            for (int i = 1; i < Sheets.Count; i++)
            {
                Sheets sheets = Workbook.Sheets;
                if (Sheets[i].IsSelected)
                {
                    var moving = sheets[i + 1];
                    var other = sheets[currentTop + 1];
                    moving.Move(before: other);
                    moving.ReleaseComObject();
                    other.ReleaseComObject();
                    Sheets.Move(i, currentTop);
                    currentTop++;
                }
                sheets.ReleaseComObject();
            }
            _lockEvents -= 1;
        }

        private bool CanMoveSheetUp()
        {
            return ((NumSelectedSheets > 0) && !Sheets[0].IsSelected);
        }

        private bool CanMoveSheetsToTop()
        {
            return CanMoveSheetUp();
        }

        private void DoMoveSheetDown()
        {
            _lockEvents += 1;
            // When iterating over the worksheet view models in the Sheets collection
            // as well as over the sheets collection of the workbook, keep in mind
            // that Excel workbook collections are 1-based.
            for (int i = Sheets.Count - 2; i >= 0; i--)
            {
                Sheets sheets = Workbook.Sheets;
                if (Sheets[i].IsSelected)
                {
                    var moving = sheets[i + 1];
                    var other = sheets[i + 2];
                    moving.Move(after: other);
                    moving.ReleaseComObject();
                    other.ReleaseComObject();
                    Sheets.Move(i, i + 1);
                }
                sheets.ReleaseComObject();
            }
            _lockEvents -= 1;
        }

        private void DoMoveSheetsToBottom()
        {
            _lockEvents += 1;
            int currentBottom = Sheets.Count - 1;
            for (int i = currentBottom-1; i >= 0; i--)
            {
                Sheets sheets = Workbook.Sheets;
                if (Sheets[i].IsSelected)
                {
                    var moving = sheets[i + 1];
                    var other = sheets[currentBottom + 1];
                    moving.Move(after: other);
                    moving.ReleaseComObject();
                    other.ReleaseComObject();
                    Sheets.Move(i, currentBottom);
                    currentBottom--;
                }
                sheets.ReleaseComObject();
            }
            _lockEvents -= 1;
        }

        private bool CanMoveSheetDown()
        {
            return ((NumSelectedSheets > 0) && !Sheets[Sheets.Count - 1].IsSelected);
        }

        private bool CanMoveSheetsToBottom()
        {
            return CanMoveSheetDown();
        }

        private void DoDeleteSheets()
        {
            ConfirmDeleteMessage.Send(
                new MessageContent(),
                (confirmation) =>
                {
                    ConfirmDeleteSheets(confirmation);
                });
        }

        private void ConfirmDeleteSheets(MessageContent confirmation)
        {
            if (confirmation.Confirmed)
            {
                Excel.ViewModels.Instance.Default.DisableDisplayAlerts();
                Sheets sheets = Workbook.Sheets;
                for (int i = 0; i < Sheets.Count; i++)
                {
                    if (Sheets[i].IsSelected)
                    {
                        // Must use sheet name rather than index in collection
                        // because indexes may differ if hidden sheets exist.
                        var s = sheets[Sheets[i].DisplayString];
                        s.Delete();
                        s.ReleaseComObject(s);
                        Sheets.RemoveAt(i);
                    }
                }
                sheets.ReleaseComObject();
                Excel.ViewModels.Instance.Default.EnableDisplayAlerts();
            }
        }

        private bool CanDeleteSheets()
        {
            return (NumSelectedSheets > 0);
        }

        private void DoRenameSheet()
        {
            Logger.Info("DoRenameSheet");
            StringMessageContent content = new StringMessageContent();
            content.Value = _lastSelectedSheet.DisplayString;
            content.Validator = (value) =>
            {
                if (SheetViewModel.IsValidName(value))
                {
                    return String.Empty;
                }
                else
                {
                    // TODO: Find a way that does not require human language here
                    return "1-21, not () /\\ [] *?";
                }
            };
            RenameSheetMessage.Send(
                content,
                (stringMessage) =>
                    {
                        ConfirmRenameSheet(stringMessage);
                    }
            );
        }

        private void ConfirmRenameSheet(StringMessageContent stringMessage)
        {
            if (CanRenameSheet() && stringMessage.Confirmed)
            {
                Logger.Info("ConfirmRenameSheet: confirmed");
                _lastSelectedSheet.DisplayString = stringMessage.Value;
            }
            else
            {
                Logger.Info("ConfirmRenameSheet: not confirmed or unable to rename sheet");
            }
        }

        private bool CanRenameSheet()
        {
            return (NumSelectedSheets > 0);
        }

        private void DoMonitorWorkbook()
        {
            if (_timer == null)
            {
                Logger.Info("Begin monitoring workbook");
                _timer = new Timer(
                    CheckSheetsChanged,
                    null,
                    XLToolbox.Properties.Settings.Default.WorkbookMonitorIntervalMilliseconds,
                    XLToolbox.Properties.Settings.Default.WorkbookMonitorIntervalMilliseconds);
            }
        }

        private bool CanMonitorWorkbook()
        {
            return _workbook != null && _timer == null;
        }

        private void DoUnmonitorWorkbook()
        {
            if (_timer != null)
            {
                Logger.Info("Stop monitoring workbook");
                _timer.Dispose();
                _timer = null;
                CheckSheetsChanged(null);
            }
        }

        private bool CanUnmonitorWorkbook()
        {
            return _timer == null;
        }

        private void CheckSheetsChanged(object state)
        {
            if (_lockEvents <= 0)
            {
                _lockEvents += 1;
                string sheetsString = SheetsString;
                if (sheetsString != _lastSheetsString)
                {
                    Logger.Info("CheckSheetsChanged: Change in worksheets detected, rebuilding list");
                    BuildSheetList();
                }
                _lockEvents -= 1;
            }
        }

        private void SheetActivated(dynamic sheet)
        {
            Logger.Debug("SheetActivated: _lockEvents is {0}", _lockEvents);
            if (sheet != null) // && _lockEvents <= 0)
            {
                _lockEvents += 1;
                SheetViewModel svm = Sheets.FirstOrDefault(s => s.IsSelected);
                if (svm != null)
                {
                    svm.IsSelected = false;
                }
                int index = IndexOf(sheet);
                Logger.Info("SheetActivated: Sheet index is {0}", index);
                if (index >= 0 && index < Sheets.Count)
                {
                    Sheets[index].IsSelected = true;
                }
                else
                {
                    Logger.Warn("SheetActivated: Index {0} is out of bounds!", index);
                }
                _lockEvents -= 1;
            }
            Dispatch(() => OnPropertyChanged("ActiveSheet"));
        }

        private void CopyProperty(object param)
        {
            if (CanCopyProperty(null))
            {
                Logger.Info("CopyProperty: Copying property to clipboard, index = {0}",
                    PreferredPropertyIndex);
                Properties[PreferredPropertyIndex].Copy();
                CloseViewCommand.Execute(null);
            }
            else
            {
                Logger.Warn("CopyProperty: Cannot copy property");
                if (Properties != null)
                {
                    Logger.Warn("CopyProperty: Properties.Count = {0}, PreferredPropertyIndex = {1}",
                        Properties.Count, PreferredPropertyIndex);
                }
                else
                {
                    Logger.Warn("CopyProperty: Properties is null");
                }
            }
        }

        private bool CanCopyProperty(object param)
        {
            return (Properties != null) && (PreferredPropertyIndex < Properties.Count)
                && (PreferredPropertyIndex >= 0);
        }

        #endregion

        #region Private fields

        private Workbook _workbook;
        private ObservableCollection<SheetViewModel> _sheets;
        private SheetViewModel _lastSelectedSheet;
        private DelegatingCommand _moveSheetUp;
        private DelegatingCommand _moveSheetsToTop;
        private DelegatingCommand _moveSheetDown;
        private DelegatingCommand _moveSheetsToBottom;
        private DelegatingCommand _deleteSheets;
        private DelegatingCommand _renameSheet;
        private DelegatingCommand _monitorWorkbook;
        private DelegatingCommand _unmonitorWorkbook;
        private DelegatingCommand _copyPropertyCommand;
        private Message<MessageContent> _confirmDeleteMessage;
        private Message<StringMessageContent> _renameSheetMessage;
        private string _lastSheetsString;
        private Timer _timer;
        private int _lockEvents;
        private List<WorkbookProperty> _properties;

        #endregion

        #region Protected properties

        protected Workbook Workbook
        {
            get
            {
                return _workbook;
            }
            set
            {
                bool _changed = value != _workbook;
                if (_changed && _workbook != null)
                {
                    DoUnmonitorWorkbook();
                    _workbook.SheetActivate -= SheetActivated;
                    _workbook.ReleaseComObject();
                }
                if (value == null)
                {
                    Logger.Info("Workbook_set: value is null");
                    DisplayString = String.Empty;

                    _workbook = null;
                }
                else
                {
                    Logger.Info("Workbook_set: Using new workbook");
                    _workbook = value;
                    _workbook.SheetActivate += SheetActivated;
                    DisplayString = _workbook.Name;
                }
                if (_changed)
                {
                    BuildSheetList();
                    OnPropertyChanged("Workbook");
                }
            }
        }
        
        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

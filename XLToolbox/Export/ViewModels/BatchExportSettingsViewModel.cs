/* BatchExportSettingsViewModel.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Export.Models;
using System.Diagnostics;

namespace XLToolbox.Export.ViewModels
{
    /// <summary>
    /// View model for the <see cref="Settings"/> class.
    /// </summary>
    public class BatchExportSettingsViewModel : SettingsViewModelBase
    {
        #region Factory
        /// <summary>
        /// Returns a BatchExportSettingsViewModel object that wraps
        /// the last used BatchExportSettings stored in the assembly's
        /// Properties, or Null if no stored object exists.
        /// </summary>
        /// <returns>BatchExportSettingsViewModel object with last
        /// used settings model, or Null if no such object exists.</returns>
        public static BatchExportSettingsViewModel FromLastUsed()
        {
            BatchExportSettings settings = BatchExportSettings.FromLastUsed();
            if (settings != null)
            {
                Logger.Info("FromLastUsed(): Got last used settings, creating view model");
                return new BatchExportSettingsViewModel(settings);
            }
            else
            {
                Logger.Info("FromLastUsed(): Did not get last used settings, returning null");
                return null;
            }
        }

        /// <summary>
        /// Returns a BatchExportSettingsViewModel object that wraps
        /// the last used BatchExportSettings stored in the
        /// workbookContent's hidden storage area, or the one stored
        /// in the assembly's Properties, or Null if no stored object exists.
        /// </summary>
        /// <param name="workbookContext"></param>
        /// <returns>BatchExportSettingsViewModel object with last
        /// used settings model, or Null if no such object exists.</returns>
        public static BatchExportSettingsViewModel FromLastUsed(Workbook workbookContext)
        {
            BatchExportSettings settings = BatchExportSettings.FromLastUsed(workbookContext);
            if (settings != null)
            {
                Logger.Info("FromLastUsed(workbookContext): Got settings for workbook context, creating view model");
                return new BatchExportSettingsViewModel(settings);
            }
            else
            {
                Logger.Info("FromLastUsed(workbookContext): Did not get settings for workbook context");
                return BatchExportSettingsViewModel.FromLastUsed();
            }
        }

        #endregion

        #region Public properties

        public EnumProvider<BatchExportScope> Scope
        {
            get
            {
                if (_scope == null)
                {
                    _scope = new EnumProvider<BatchExportScope>(
                        ((BatchExportSettings)Settings).Scope);
                    _scope.AsEnum = ((BatchExportSettings)Settings).Scope;
                }
                return _scope;
            }
        }

        public EnumProvider<BatchExportLayout> Layout
        {
            get
            {
                if (_layout == null)
                {
                    _layout = new EnumProvider<BatchExportLayout>(
                        ((BatchExportSettings)Settings).Layout);
                    _layout.AsEnum = ((BatchExportSettings)Settings).Layout;
                }
                return _layout;
            }
        }
        
        public EnumProvider<BatchExportObjects> Objects
        {
            get
            {
                if (_objects == null)
                {
                    _objects = new EnumProvider<BatchExportObjects>(
                        ((BatchExportSettings)Settings).Objects);
                    _objects.AsEnum = ((BatchExportSettings)Settings).Objects;
                }
                return _objects;
            }
        }

        public string Path
        {
            get { return ((BatchExportSettings)Settings).Path; }
            set
            {
                ((BatchExportSettings)Settings).Path = value;
                OnPropertyChanged("Path");
            }
        }

        public bool IsActiveSheetEnabled
        {
            get
            {
                return _isActiveSheetEnabled;
            }
            set
            {
                _isActiveSheetEnabled = value;
                OnPropertyChanged("IsActiveSheetEnabled");
            }
        }

        public bool IsActiveWorkbookEnabled
        {
            get
            {
                return _isActiveWorkbookEnabled;
            }
            set
            {
                _isActiveWorkbookEnabled = value;
                OnPropertyChanged("IsAllSheetsEnabled");
            }
        }

        public bool IsOpenWorkbooksEnabled
        {
            get
            {
                return _isOpenWorkbooksEnabled;
            }
            set
            {
                _isOpenWorkbooksEnabled = value;
                if (!value && Scope.AsEnum == BatchExportScope.OpenWorkbooks) Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                OnPropertyChanged("IsAllWorkbooksEnabled");
            }
        }

        public bool IsChartsEnabled
        {
            get
            {
                return _isChartsEnabled;
            }
            set
            {
                _isChartsEnabled = value;
                OnPropertyChanged("IsChartsEnabled");
            }
        }

        public bool IsChartsAndShapesEnabled
        {
            get
            {
                return _isChartsAndShapesEnabled;
            }
            set
            {
                _isChartsAndShapesEnabled = value;
                if (!value && Objects.AsEnum == BatchExportObjects.ChartsAndShapes) Objects.AsEnum = BatchExportObjects.Charts;
                OnPropertyChanged("IsChartsAndShapesEnabled");
            }
        }

        public bool IsSingleItemsEnabled
        {
            get
            {
                return _isSingleItemsEnabled;
            }
            set
            {
                _isSingleItemsEnabled = value;
                OnPropertyChanged("IsSingleItemsEnabled");
            }
        }

        public bool IsSheetLayoutEnabled
        {
            get
            {
                return _isSheetLayoutEnabled;
            }
            set
            {
                _isSheetLayoutEnabled = value;
                if (!value && Layout.AsEnum == BatchExportLayout.SheetLayout) Layout.AsEnum = BatchExportLayout.SingleItems;
                OnPropertyChanged("IsSheetLayoutEnabled");
            }
        }
       
        #endregion

        #region Public methods

        /// <summary>
        /// Makes sure that no disabled options are selected.
        /// </summary>
        /// <remarks>
        /// This method is not automatically called in the constructor
        /// in order to allow subscribed views to decide whether to
        /// 'sanitize' or not. For example, the
        /// <see cref="QuickExporter.ExportBatch()"/> method deliberately 
        /// refrains from 'sanitizing' so that the user can see what
        /// options are selected, but disabled.
        /// </remarks>
        public void SanitizeOptions()
        {
            Logger.Info("SanitizeOptions");
            if (!CanExport()) {
                if ((Scope.AsEnum == BatchExportScope.ActiveSheet) && !IsActiveSheetEnabled)
                    Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                if ((Scope.AsEnum == BatchExportScope.ActiveWorkbook) && !IsActiveWorkbookEnabled)
                    Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                if ((Scope.AsEnum == BatchExportScope.OpenWorkbooks) && !IsOpenWorkbooksEnabled)
                    Scope.AsEnum = BatchExportScope.ActiveSheet;

                if ((Objects.AsEnum == BatchExportObjects.Charts) && !IsChartsEnabled)
                    Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                if ((Objects.AsEnum == BatchExportObjects.ChartsAndShapes) && !IsChartsAndShapesEnabled)
                    Objects.AsEnum = BatchExportObjects.Charts;

                if ((Layout.AsEnum == BatchExportLayout.SingleItems) && !IsSingleItemsEnabled)
                    Layout.AsEnum = BatchExportLayout.SheetLayout;
                if ((Layout.AsEnum == BatchExportLayout.SheetLayout) && !IsSheetLayoutEnabled)
                    Layout.AsEnum = BatchExportLayout.SingleItems;
            }
        }

        #endregion

        #region Commands

        /// <summary>
        /// Causes the <see cref="ChooseFolderMessage"/> to be sent.
        /// Upon confirmation of the message by a view, the Export
        /// process will be started.
        /// </summary>
        public DelegatingCommand ChooseFolderCommand
        {
            get
            {
                if (_chooseFolderCommand == null)
                {
                    _chooseFolderCommand = new DelegatingCommand(
                        param => DoChooseFolder(),
                        param => CanChooseFolder());
                }
                return _chooseFolderCommand;
            }
        }

        #endregion

        #region Messages

        public Message<FileNameMessageContent> ChooseFolderMessage
        {
            get
            {
                if (_chooseFolderMessage == null)
                {
                    _chooseFolderMessage = new Message<FileNameMessageContent>();
                }
                return _chooseFolderMessage;
            }
        }

        #endregion

        #region Constructors

        public BatchExportSettingsViewModel()
            : this(new BatchExportSettings())
        { }

        public BatchExportSettingsViewModel(BatchExportSettings settings)
            : this(new BatchExporter(settings as BatchExportSettings))
        { }

        public BatchExportSettingsViewModel(BatchExporter batchExporter)
            : base(batchExporter)
        {
            Settings = batchExporter.Settings;
            PresetViewModels.Select(Settings.Preset);
            if (String.IsNullOrEmpty(FileName))
            {
                FileName = String.Format("{{{0}}}_{{{1}}}_{{{2}}}",
                    Strings.Workbook, Strings.Worksheet, Strings.Index);
            }
            Scope.PropertyChanged += Scope_PropertyChanged;
            Objects.PropertyChanged += Objects_PropertyChanged;
            Layout.PropertyChanged += Layout_PropertyChanged;
            UpdateStates();
        }

        #endregion

        #region Implementation of abstract methods
        
        /// <summary>
        /// Determines the suggested target directory and sends the
        /// ChooseFileNameMessage.
        /// </summary>
        private void DoChooseFolder()
        {
            Logger.Info("DoChooseFolder");
            string path = ((BatchExportSettings)Settings).Path;
            if (string.IsNullOrEmpty(path))
            {
                path = LoadExportPath();
            }
            ChooseFolderMessage.Send(
                new FileNameMessageContent(path, null),
                (content) => ConfirmFolder(content)
            );
        }

        private bool CanChooseFolder()
        {
            return CanExport();
        }

        protected override void UpdateProcessMessageContent(ProcessMessageContent processMessageContent)
        {
            if (Exporter != null)
            {
                processMessageContent.PercentCompleted = Exporter.PercentCompleted;
            }
            else
            {
                Logger.Warn("UpdateProcessMessageContent: Exporter is null!");
            }
        }

        protected override void DoExport()
        {
            Logger.Info("DoExport");
            if (CanExport())
            {
                ((BatchExportSettings)Settings).Store(Instance.Default.ActiveWorkbook);
                StartProcess();
            }
        }

        protected override bool CanExport()
        {
            return (Settings != null) &&
                (Settings.Preset != null) &&
                CanExecuteMatrix[Scope.AsEnum][Objects.AsEnum][Layout.AsEnum];
        }

        #endregion

        #region Event handlers

        private void Scope_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            ((BatchExportSettings)Settings).Scope = Scope.AsEnum;
            SetObjectsState();
            SetLayoutState();
        }

        private void Objects_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            ((BatchExportSettings)Settings).Objects = Objects.AsEnum;
            SetLayoutState();
        }

        private void Layout_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            ((BatchExportSettings)Settings).Layout = Layout.AsEnum;
        }

        #endregion

        #region Private methods

        private void UpdateStates()
        {
            Logger.Info("UpdateStates");
            AnalyzeOpenWorkbooks();
            SetScopeState();
            SetObjectsState();
            SetLayoutState();
        }

        private void AnalyzeOpenWorkbooks()
        {
            Logger.Info("AnalyzeOpenWorkboks");
            AnalyzeActiveSheet();
            AnalyzeOtherSheets();
            AnalyzeOtherWorkbooks();
        }

        private void AnalyzeActiveSheet()
        {
            if (Instance.Default.Application.ActiveSheet != null)
            {
                Logger.Info("AnalyzeActiveSheet");
                SheetViewModel sheetVM = new SheetViewModel(Instance.Default.Application.ActiveSheet);
                int charts = sheetVM.CountCharts();
                int shapes = sheetVM.CountShapes() - charts;
                _activeSheetHasCharts = charts > 0;
                _activeSheetHasManyCharts = charts > 1;
                _activeSheetHasShapes = shapes > 0;
                _activeSheetHasManyObjects = charts + shapes > 1;
                _activeSheetName = sheetVM.DisplayString;
            }
            else
            {
                Logger.Info("AnalyzeActiveSheet: No active sheet");
            }
        }

        private void AnalyzeOtherSheets()
        {
            Logger.Info("AnalyzeOtherSheets");
            _otherSheetsHaveCharts = false;
            _otherSheetsHaveManyCharts = false;
            _otherSheetsHaveShapes = false;
            _otherSheetsHaveManyObjects = false;
            int charts;
            int shapes;
            object sheet;
            Sheets sheets = Instance.Default.ActiveWorkbook.Sheets;
            Logger.Info("AnalyzeOtherSheets: {0} sheet(s) in active workbook", sheets.Count);
            for (int i = 1; i <= sheets.Count; i++)
            {
                sheet = sheets[i];
                SheetViewModel svm = new SheetViewModel(sheet);
                if (svm.DisplayString != _activeSheetName)
                {
                    charts = svm.CountCharts();
                    shapes = svm.CountShapes() - charts;
                    Logger.Info("AnalyzeOtherSheets: [{0}]: charts: {1}, shapes: {2}", i, charts, shapes);
                    _otherSheetsHaveCharts |= charts > 0;
                    _otherSheetsHaveManyCharts |= charts > 1;
                    _otherSheetsHaveShapes |= shapes > 0;
                    _otherSheetsHaveManyObjects |= charts + shapes > 1;
                }
                else
                {
                    Logger.Info("AnalyzeOtherSheets: [{0}]: is active sheet", i);
                }
                Bovender.ComHelpers.ReleaseComObject(sheet);
            }
            Logger.Info("AnalyzeOtherSheets: Releasing sheets object");
            Bovender.ComHelpers.ReleaseComObject(sheets);
        }

        private void AnalyzeOtherWorkbooks()
        {
            Logger.Info("AnalyzeOtherWorkbooks");
            string activeWorkbookName = Instance.Default.ActiveWorkbook.Name;
            _otherWorkbooksHaveCharts = false;
            _otherWorkbooksHaveManyCharts = false;
            _otherWorkbooksHaveShapes = false;
            _otherWorkbooksHaveManyObjects = false;
            int charts;
            int shapes;
            Workbooks workbooks = Instance.Default.Application.Workbooks;
            Logger.Info("AnalyzeOtherWorkbooks: {0} workbook(s) are currently open", workbooks.Count);
            for (int i = 1; i <= workbooks.Count; i++)
            {
                Logger.Info("AnalyzeOtherWorkbooks: [{0}]", i);
                Workbook workbook = workbooks[i];
                if (workbook.Name != activeWorkbookName)
                {
                    Sheets sheets = workbook.Sheets;
                    for (int j = 1; j <= sheets.Count; j++)
                    {
                        object sheet = sheets[j];
                        SheetViewModel svm = new SheetViewModel(sheet);
                        charts = svm.CountCharts();
                        shapes = svm.CountShapes() - charts;
                        _otherWorkbooksHaveCharts |= charts > 0;
                        _otherWorkbooksHaveManyCharts |= charts > 1;
                        _otherWorkbooksHaveShapes |= shapes > 0;
                        _otherWorkbooksHaveManyObjects |= charts + shapes > 1;
                        Bovender.ComHelpers.ReleaseComObject(sheet);
                    }
                    Bovender.ComHelpers.ReleaseComObject(sheets);
                }
                else
                {
                    Logger.Info("AnalyzeOtherWorkbooks: [{0}] is the active workbook", i);
                }
                Bovender.ComHelpers.ReleaseComObject(workbook);
            }
            Logger.Info("AnalyzeOtherWorkbooks: Releasing workbooks object");
            Bovender.ComHelpers.ReleaseComObject(workbooks);
        }

        private void FillCanExecuteMatrix()
        {
            _canExecuteMatrix = new ScopeStates()
            {
                { BatchExportScope.ActiveSheet, new ObjectsStates()
                    {
                        { BatchExportObjects.Charts, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, _activeSheetHasCharts },
                                { BatchExportLayout.SheetLayout, _activeSheetHasManyCharts }
                            }
                        },
                        { BatchExportObjects.ChartsAndShapes, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems,
                                    _activeSheetHasCharts | _activeSheetHasShapes },
                                { BatchExportLayout.SheetLayout,
                                    _activeSheetHasManyObjects }
                            }
                        },
                    }
                },
                { BatchExportScope.ActiveWorkbook, new ObjectsStates()
                    {
                        { BatchExportObjects.Charts, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, 
                                    _activeSheetHasCharts | _otherSheetsHaveCharts },
                                { BatchExportLayout.SheetLayout, 
                                    _activeSheetHasManyCharts | _otherSheetsHaveManyCharts }
                            }
                        },
                        { BatchExportObjects.ChartsAndShapes, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, 
                                    _activeSheetHasCharts | _otherSheetsHaveCharts |
                                    _activeSheetHasShapes | _otherSheetsHaveShapes },
                                { BatchExportLayout.SheetLayout, 
                                    _activeSheetHasManyObjects | _otherSheetsHaveManyObjects }
                            }
                        },
                    }
                },
                { BatchExportScope.OpenWorkbooks, new ObjectsStates()
                    {
                        { BatchExportObjects.Charts, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, 
                                    _activeSheetHasCharts | _otherSheetsHaveCharts |
                                    _otherWorkbooksHaveCharts },
                                { BatchExportLayout.SheetLayout, 
                                    _activeSheetHasManyCharts | _otherSheetsHaveManyCharts |
                                    _otherWorkbooksHaveManyCharts }
                            }
                        },
                        { BatchExportObjects.ChartsAndShapes, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, 
                                    _activeSheetHasCharts | _otherSheetsHaveCharts |
                                    _activeSheetHasShapes | _otherSheetsHaveShapes |
                                    _otherWorkbooksHaveCharts | _otherWorkbooksHaveShapes },
                                { BatchExportLayout.SheetLayout, 
                                    _activeSheetHasManyObjects | _otherSheetsHaveManyObjects |
                                    _otherWorkbooksHaveManyObjects }
                            }
                        },
                    }
                },
            };
        }

        /// <summary>
        /// Sets the enabled state of the Scope property. For performance
        /// reasons, this is only done then the class is instantiated;
        /// if the state of the Excel application changes, this will
        /// not be noticed by this class.
        /// </summary>
        private void SetScopeState()
        {
            IsActiveSheetEnabled = _activeSheetHasCharts | _activeSheetHasShapes;
            IsActiveWorkbookEnabled = _otherSheetsHaveCharts | _otherSheetsHaveShapes;
            IsOpenWorkbooksEnabled = _otherWorkbooksHaveCharts | _otherWorkbooksHaveShapes;
        }

        /// <summary>
        /// Sets the enabled state of the Objects property, depending
        /// on the current value of the Scope property.
        /// </summary>
        private void SetObjectsState()
        {
            switch (Scope.AsEnum)
            {
                case BatchExportScope.ActiveSheet:
                    IsChartsEnabled = _activeSheetHasCharts;
                    IsChartsAndShapesEnabled = _activeSheetHasShapes;
                    break;
                case BatchExportScope.ActiveWorkbook:
                    IsChartsEnabled = _activeSheetHasCharts | _otherSheetsHaveCharts;
                    IsChartsAndShapesEnabled = _activeSheetHasShapes | _otherSheetsHaveShapes;
                    break;
                case BatchExportScope.OpenWorkbooks:
                    IsChartsEnabled = _activeSheetHasCharts | _otherSheetsHaveCharts |
                        _otherWorkbooksHaveCharts;
                    IsChartsAndShapesEnabled = _activeSheetHasShapes | _otherSheetsHaveShapes |
                        _otherWorkbooksHaveShapes;
                    break;
                default:
                    throw new InvalidOperationException(
                        "No case defined for " + Scope.SelectedItem);
            }
        }

        /// <summary>
        /// Sets the enabled state of the Layout property, depending
        /// on the current value of the Scope and Objects properties.
        /// </summary>
        private void SetLayoutState()
        {
            switch (Scope.AsEnum)
            {
                case BatchExportScope.ActiveSheet:
                    switch (Objects.AsEnum)
                    {
                        case BatchExportObjects.Charts:
                            IsSingleItemsEnabled = _activeSheetHasCharts;
                            IsSheetLayoutEnabled = _activeSheetHasManyCharts;
                            break;
                        case BatchExportObjects.ChartsAndShapes:
                            IsSingleItemsEnabled = _activeSheetHasCharts | _activeSheetHasShapes;
                            IsSheetLayoutEnabled = _activeSheetHasManyObjects;
                            break;
                        default:
                            throw new InvalidOperationException(String.Format(
                                "No case for {0} and {1}", Scope.SelectedItem, Objects.SelectedItem));
                    }
                    break;
                case BatchExportScope.ActiveWorkbook:
                    switch (Objects.AsEnum)
                    {
                        case BatchExportObjects.Charts:
                            IsSingleItemsEnabled = _activeSheetHasCharts | _otherSheetsHaveCharts;
                            IsSheetLayoutEnabled = _activeSheetHasManyCharts | _otherSheetsHaveManyCharts;
                            break;
                        case BatchExportObjects.ChartsAndShapes:
                            IsSingleItemsEnabled =
                                _activeSheetHasCharts | _otherSheetsHaveCharts |
                                _activeSheetHasShapes | _otherSheetsHaveShapes;
                            IsSheetLayoutEnabled =
                                _activeSheetHasManyObjects | _otherSheetsHaveManyObjects;
                            break;
                        default:
                            throw new InvalidOperationException(String.Format(
                                "No case for {0} and {1}", Scope.SelectedItem, Objects.SelectedItem));
                    }
                    break;
                case BatchExportScope.OpenWorkbooks:
                    switch (Objects.AsEnum)
                    {
                        case BatchExportObjects.Charts:
                            IsSingleItemsEnabled =
                                _activeSheetHasCharts | _otherSheetsHaveCharts | _otherWorkbooksHaveCharts;
                            IsSheetLayoutEnabled =
                                _activeSheetHasManyCharts | _otherSheetsHaveManyCharts | _otherWorkbooksHaveManyCharts;
                            break;
                        case BatchExportObjects.ChartsAndShapes:
                            IsSingleItemsEnabled =
                                _activeSheetHasCharts | _otherSheetsHaveCharts |
                                _activeSheetHasShapes | _otherSheetsHaveShapes |
                                _otherWorkbooksHaveCharts | _otherWorkbooksHaveShapes;
                            IsSheetLayoutEnabled =
                                _activeSheetHasManyObjects | _otherSheetsHaveManyObjects |
                                _otherWorkbooksHaveManyObjects;
                            break;
                        default:
                            throw new InvalidOperationException(String.Format(
                                "No case for {0} and {1}", Scope.SelectedItem, Objects.SelectedItem));
                    }
                    break;
                default:
                    throw new InvalidOperationException(
                        "No case defined for " + Scope.SelectedItem);
            }
        }

        private void ConfirmFolder(FileNameMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                Logger.Info("ConfirmFolder: Confirmed");
                ((BatchExportSettings)Settings).Path = messageContent.Value;
                DoExport();
            }
            else
            {
                Logger.Info("ConfirmFolder: Not confirmed");
            }
        }

        private bool CanExportFromSheet(object sheet)
        {
            SheetViewModel svm = new SheetViewModel(
                Instance.Default.Application.ActiveSheet);
            switch (Objects.AsEnum)
            {
                case BatchExportObjects.Charts:
                    return svm.CountCharts() > 0;
                case BatchExportObjects.ChartsAndShapes:
                    return svm.CountShapes() > 0;
                default:
                    Logger.Fatal("CanExportFromSheet: Unknown case: {0}", Objects.AsEnum);
                    throw new InvalidOperationException(
                        "Cannot handle " + Objects.SelectedItem);
            }
        }

        private bool CanExportFromWorkbook(Workbook workbook)
        {
            bool result = false;
            Sheets sheets = workbook.Sheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                object sheet = sheets[i];
                if (CanExportFromSheet(sheet)) result = true;
                Bovender.ComHelpers.ReleaseComObject(sheet);
                if (result) break;
            }
            Logger.Info("CanExportFromWorkbook: {0}", result);
            Bovender.ComHelpers.ReleaseComObject(sheets);
            return result;
        }

        #endregion

        #region Private properties

        private ScopeStates CanExecuteMatrix
        {
            get
            {
                if (_canExecuteMatrix == null)
                {
                    FillCanExecuteMatrix();
                }
                return _canExecuteMatrix;
            }
        }

        private BatchExporter Exporter
        {
            [DebuggerStepThrough]
            get
            {
                return ProcessModel as BatchExporter;
            }
        }

        #endregion

        #region Private fields

        private EnumProvider<BatchExportScope> _scope;
        private EnumProvider<BatchExportObjects> _objects;
        private EnumProvider<BatchExportLayout> _layout;
        private DelegatingCommand _chooseFolderCommand;
        private Message<FileNameMessageContent> _chooseFolderMessage;

        private bool _isActiveSheetEnabled;
        private bool _isActiveWorkbookEnabled;
        private bool _isOpenWorkbooksEnabled;
        private bool _isChartsEnabled;
        private bool _isChartsAndShapesEnabled;
        private bool _isSingleItemsEnabled;
        private bool _isSheetLayoutEnabled;

        private bool _activeSheetHasCharts;
        private bool _activeSheetHasShapes;
        private bool _activeSheetHasManyCharts;
        private bool _activeSheetHasManyObjects;

        private bool _otherSheetsHaveCharts;
        private bool _otherSheetsHaveShapes;
        private bool _otherSheetsHaveManyCharts;
        private bool _otherSheetsHaveManyObjects;

        private bool _otherWorkbooksHaveCharts;
        private bool _otherWorkbooksHaveShapes;
        private bool _otherWorkbooksHaveManyCharts;
        private bool _otherWorkbooksHaveManyObjects;

        private string _activeSheetName;

        private ScopeStates _canExecuteMatrix;

        #endregion

        #region Embedded classes

        class LayoutStates : Dictionary<BatchExportLayout, bool> { }
        class ObjectsStates : Dictionary<BatchExportObjects, LayoutStates> { }
        class ScopeStates : Dictionary<BatchExportScope, ObjectsStates> { }

        #endregion

        #region Class logger

        new private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

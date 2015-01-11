/* BatchExportSettingsViewModel.cs
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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Excel.Instance;
using XLToolbox.WorkbookStorage;
using XLToolbox.Export.Models;

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
                return new BatchExportSettingsViewModel(settings);
            }
            else
            {
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
                return new BatchExportSettingsViewModel(settings);
            }
            else
            {
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
                OnPropertyChanged("IsPreserveLayoutEnabled");
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

        public Message<StringMessageContent> ChooseFolderMessage
        {
            get
            {
                if (_chooseFolderMessage == null)
                {
                    _chooseFolderMessage = new Message<StringMessageContent>();
                }
                return _chooseFolderMessage;
            }
        }

        #endregion

        #region Constructors

        public BatchExportSettingsViewModel()
            : this(new BatchExportSettings())
        {
            if (PresetsRepository.Presets.Count > 0)
            {
                PresetsRepository.Presets[0].IsSelected = true;
            }
        }

        public BatchExportSettingsViewModel(BatchExportSettings settings)
            : base()
        {
            Settings = settings;
            if (settings.Preset != null)
            {
                PresetViewModel pvm = new PresetViewModel(settings.Preset);
                if (!PresetsRepository.Select(pvm))
                {
                    PresetsRepository.Presets.Add(pvm);
                    pvm.IsSelected = true;
                }
            }
            FileName = String.Format("{{{0}}}_{{{1}}}_{{{2}}}",
                Strings.Workbook, Strings.Worksheet, Strings.Index);
            Scope.PropertyChanged += Scope_PropertyChanged;
            Objects.PropertyChanged += Objects_PropertyChanged;
            Layout.PropertyChanged += Layout_PropertyChanged;
            UpdateStates();
        }

        #endregion

        #region Implementation of SettingsViewModelBase

        /// <summary>
        /// Determines the suggested target directory and sends the
        /// ChooseFileNameMessage.
        /// </summary>
        private void DoChooseFolder()
        {
            string path = ((BatchExportSettings)Settings).Path;
            if (string.IsNullOrEmpty(path))
            {
                path = GetExportPath();
            }
            ChooseFolderMessage.Send(
                new StringMessageContent(path),
                (content) => ConfirmFolder(content)
            );
        }

        private bool CanChooseFolder()
        {
            return CanExport();
        }

        protected override void DoExport()
        {
            if (CanExport())
            {
                ((BatchExportSettings)Settings).Store(
                    ExcelInstance.Application.ActiveWorkbook);
                SaveExportPath();
                Exporter exporter = new Exporter();
                ProcessMessageContent processMessageContent =
                    new ProcessMessageContent(exporter.CancelBatchExport);
                exporter.BatchExportProgressChanged +=
                    (object sender, ExportProgressChangedEventArgs args) =>
                    {
                        processMessageContent.PercentCompleted = args.PercentCompleted;
                    };
                exporter.BatchExportFinished +=
                    (object sender, ExportFinishedEventArgs args) =>
                    {
                        Dispatcher.Invoke(new System.Action(
                                () =>
                                {
                                    processMessageContent.CompletedMessage.Send(processMessageContent);
                                }
                            )
                        );
                    };
                processMessageContent.Processing = true;
                ExportProcessMessage.Send(processMessageContent);
                exporter.ExportBatchAsync(Settings as BatchExportSettings);
            }
        }

        protected override bool CanExport()
        {
            return CanExecuteMatrix[Scope.AsEnum][Objects.AsEnum][Layout.AsEnum] &&
                (SelectedPreset != null);
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
            AnalyzeOpenWorkbooks();
            SetScopeState();
            SetObjectsState();
            SetLayoutState();
        }

        private void AnalyzeOpenWorkbooks()
        {
            if (ExcelInstance.Running)
            {
                AnalyzeActiveSheet();
                AnalyzeOtherSheets();
                AnalyzeOtherWorkbooks();
            }
        }

        private void AnalyzeActiveSheet()
        {
            if (ExcelInstance.Application.ActiveSheet != null)
            {
                SheetViewModel sheetVM = new SheetViewModel(ExcelInstance.Application.ActiveSheet);
                int charts = sheetVM.CountCharts();
                int shapes = sheetVM.CountShapes() - charts;
                _activeSheetHasCharts = charts > 0;
                _activeSheetHasManyCharts = charts > 1;
                _activeSheetHasShapes = shapes > 0;
                _activeSheetHasManyObjects = charts + shapes > 1;
                _activeSheetName = sheetVM.DisplayString;
            }
        }

        private void AnalyzeOtherSheets()
        {
            _otherSheetsHaveCharts = false;
            _otherSheetsHaveManyCharts = false;
            _otherSheetsHaveShapes = false;
            _otherSheetsHaveManyObjects = false;
            int charts;
            int shapes;
            foreach (object sheet in ExcelInstance.Application.ActiveWorkbook.Sheets)
            {
                SheetViewModel svm = new SheetViewModel(sheet);
                if (svm.DisplayString != _activeSheetName)
                {
                    charts = svm.CountCharts();
                    shapes = svm.CountShapes() - charts;
                    _otherSheetsHaveCharts |= charts > 0;
                    _otherSheetsHaveManyCharts |= charts > 1;
                    _otherSheetsHaveShapes |= shapes > 0;
                    _otherSheetsHaveManyObjects |= charts + shapes > 1;
                }
            }
        }

        private void AnalyzeOtherWorkbooks()
        {
            string activeWorkbookName = ExcelInstance.Application.ActiveWorkbook.Name;
            _otherWorkbooksHaveCharts = false;
            _otherWorkbooksHaveManyCharts = false;
            _otherWorkbooksHaveShapes = false;
            _otherWorkbooksHaveManyObjects = false;
            int charts;
            int shapes;
            foreach (Workbook workbook in ExcelInstance.Application.Workbooks)
            {
                if (workbook.Name != activeWorkbookName)
                {
                    foreach (object sheet in workbook.Sheets)
                    {
                        SheetViewModel svm = new SheetViewModel(sheet);
                        charts = svm.CountCharts();
                        shapes = svm.CountShapes() - charts;
                        _otherWorkbooksHaveCharts |= charts > 0;
                        _otherWorkbooksHaveManyCharts |= charts > 1;
                        _otherWorkbooksHaveShapes |= shapes > 0;
                        _otherWorkbooksHaveManyObjects |= charts + shapes > 1;
                    }
                }
            }
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

        private void ConfirmFolder(StringMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                ((BatchExportSettings)Settings).Path = messageContent.Value;
                DoExport();
            }
        }

        private bool CanExportFromSheet(object sheet)
        {
            SheetViewModel svm = new SheetViewModel(
                ExcelInstance.Application.ActiveSheet);
            switch (Objects.AsEnum)
            {
                case BatchExportObjects.Charts:
                    return svm.CountCharts() > 0;
                case BatchExportObjects.ChartsAndShapes:
                    return svm.CountShapes() > 0;
                default:
                    throw new InvalidOperationException(
                        "Cannot handle " + Objects.SelectedItem);
            }
        }

        private bool CanExportFromWorkbook(Workbook workbook)
        {
            foreach (object sheet in workbook.Sheets)
            {
                if (CanExportFromSheet(sheet)) return true;
            }
            return false;
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

        #endregion

        #region Private fields

        private EnumProvider<BatchExportScope> _scope;
        private EnumProvider<BatchExportObjects> _objects;
        private EnumProvider<BatchExportLayout> _layout;
        private DelegatingCommand _chooseFolderCommand;
        private Message<StringMessageContent> _chooseFolderMessage;

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
    }
}

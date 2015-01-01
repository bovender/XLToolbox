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
            : base()
        {
            Settings = new BatchExportSettings();
            FileName = String.Format("{{{0}}}_{{{1}}}_{{{2}}}",
                Strings.Workbook, Strings.Worksheet, Strings.Index);
            Scope.PropertyChanged += Scope_PropertyChanged;
            Objects.PropertyChanged += Objects_PropertyChanged;
            Layout.PropertyChanged += Layout_PropertyChanged;
            UpdateStates();
        }

        public BatchExportSettingsViewModel(BatchExportSettings settings)
            : this()
        {
            Settings = settings;
            PresetViewModel pvm = new PresetViewModel(settings.Preset);
            if (!PresetsRepository.Select(pvm))
            {
                PresetsRepository.Presets.Add(pvm);
                pvm.IsSelected = true;
            }
        }

        #endregion

        #region Implementation of SettingsViewModelBase

        /// <summary>
        /// Determines the suggested target directory and sends the
        /// ChooseFileNameMessage.
        /// </summary>
        private void DoChooseFolder()
        {
            ChooseFolderMessage.Send(
                new StringMessageContent(GetExportPath()),
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
                // TODO: Make export asynchronous
                ProcessMessageContent pcm = new ProcessMessageContent();
                ExportProcessMessage.Send(pcm);
                Exporter exporter = new Exporter();
                exporter.ExportBatch(Settings as BatchExportSettings);
                pcm.CompletedMessage.Send(pcm);
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
                _numChartsActiveSheet = sheetVM.CountCharts();
                _numShapesActiveSheet = sheetVM.CountShapes() - _numChartsActiveSheet;
                _activeSheetName = sheetVM.DisplayString;
            }
        }

        private void AnalyzeOtherSheets()
        {
            foreach (object sheet in ExcelInstance.Application.ActiveWorkbook.Sheets)
            {
                SheetViewModel svm = new SheetViewModel(sheet);
                if (svm.DisplayString != _activeSheetName)
                {
                    _maxChartsOtherSheets = Math.Max(
                        _maxChartsOtherSheets, svm.CountCharts());
                    _maxShapesOtherSheets = Math.Max(
                        _maxShapesOtherSheets, svm.CountShapes() - svm.CountCharts());
                    _hasSheetWithMultipleChartsAndShapes |=
                        (svm.CountShapes() > 1) && (svm.CountShapes() > svm.CountCharts());
                }
            }
        }

        private void AnalyzeOtherWorkbooks()
        {
            string activeWorkbookName = ExcelInstance.Application.ActiveWorkbook.Name;
            foreach (Workbook workbook in ExcelInstance.Application.Workbooks)
            {
                if (workbook.Name != activeWorkbookName)
                {
                    foreach (object sheet in workbook.Sheets)
                    {
                        SheetViewModel svm = new SheetViewModel(sheet);
                        _maxChartsOtherWorkbooks = Math.Max(
                            _maxChartsOtherWorkbooks, svm.CountCharts());
                        _maxShapesOtherWorkbooks = Math.Max(
                            _maxShapesOtherWorkbooks, svm.CountShapes() - svm.CountCharts());
                        _anyWorkbookMultipleChartsAndShapes |=
                            (svm.CountShapes() > 1) && (svm.CountShapes() > svm.CountCharts());
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
                                { BatchExportLayout.SingleItems, _numChartsActiveSheet > 0 },
                                { BatchExportLayout.SheetLayout, _numChartsActiveSheet > 1 }
                            }
                        },
                        { BatchExportObjects.ChartsAndShapes, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, _numShapesActiveSheet > 0 },
                                { BatchExportLayout.SheetLayout,
                                    (_numShapesActiveSheet + _numChartsActiveSheet > 1) &&
                                    (_numShapesActiveSheet > 0) }
                            }
                        },
                    }
                },
                { BatchExportScope.ActiveWorkbook, new ObjectsStates()
                    {
                        { BatchExportObjects.Charts, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, _maxChartsOtherSheets > 0 },
                                { BatchExportLayout.SheetLayout, _maxChartsOtherSheets > 1 }
                            }
                        },
                        { BatchExportObjects.ChartsAndShapes, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, _maxShapesOtherSheets > 0 },
                                { BatchExportLayout.SheetLayout, _hasSheetWithMultipleChartsAndShapes }
                            }
                        },
                    }
                },
                { BatchExportScope.OpenWorkbooks, new ObjectsStates()
                    {
                        { BatchExportObjects.Charts, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, _maxChartsOtherWorkbooks > 0 },
                                { BatchExportLayout.SheetLayout, _maxChartsOtherWorkbooks > 1 }
                            }
                        },
                        { BatchExportObjects.ChartsAndShapes, new LayoutStates()
                            {
                                { BatchExportLayout.SingleItems, _maxShapesOtherWorkbooks > 0 },
                                { BatchExportLayout.SheetLayout, _anyWorkbookMultipleChartsAndShapes }
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
            IsActiveSheetEnabled = _numChartsActiveSheet + _numShapesActiveSheet > 0;
            IsActiveWorkbookEnabled = _maxChartsOtherSheets + _maxShapesOtherSheets > 0;
            IsOpenWorkbooksEnabled = _maxChartsOtherWorkbooks + _maxShapesOtherWorkbooks > 0;
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
                    IsChartsEnabled = _numChartsActiveSheet > 0;
                    IsChartsAndShapesEnabled = _numShapesActiveSheet > 0;
                    break;
                case BatchExportScope.ActiveWorkbook:
                    IsChartsEnabled = _maxChartsOtherSheets > 0;
                    IsChartsAndShapesEnabled = _maxShapesOtherSheets > 0;
                    break;
                case BatchExportScope.OpenWorkbooks:
                    IsChartsEnabled = _maxChartsOtherWorkbooks > 0;
                    IsChartsAndShapesEnabled = _maxShapesOtherWorkbooks > 0;
                    break;
                default:
                    throw new InvalidOperationException(
                        "No case defined for " + Scope.AsString);
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
                            IsSingleItemsEnabled = _numChartsActiveSheet > 0;
                            IsSheetLayoutEnabled = _numChartsActiveSheet > 1;
                            break;
                        case BatchExportObjects.ChartsAndShapes:
                            IsSingleItemsEnabled = _numShapesActiveSheet > 0;
                            IsSheetLayoutEnabled =
                                (_numChartsActiveSheet + _numShapesActiveSheet > 1) &&
                                (_numShapesActiveSheet > 0);
                            break;
                        default:
                            throw new InvalidOperationException(String.Format(
                                "No case for {0} and {1}", Scope.AsString, Objects.AsString));
                    }
                    break;
                case BatchExportScope.ActiveWorkbook:
                    switch (Objects.AsEnum)
                    {
                        case BatchExportObjects.Charts:
                            IsSingleItemsEnabled = _maxChartsOtherSheets > 0;
                            IsSheetLayoutEnabled = _maxChartsOtherSheets > 1;
                            break;
                        case BatchExportObjects.ChartsAndShapes:
                            IsSingleItemsEnabled = _maxShapesOtherSheets > 0;
                            IsSheetLayoutEnabled = _hasSheetWithMultipleChartsAndShapes;
                            break;
                        default:
                            throw new InvalidOperationException(String.Format(
                                "No case for {0} and {1}", Scope.AsString, Objects.AsString));
                    }
                    break;
                case BatchExportScope.OpenWorkbooks:
                    switch (Objects.AsEnum)
                    {
                        case BatchExportObjects.Charts:
                            IsSingleItemsEnabled = _maxChartsOtherWorkbooks > 0;
                            IsSheetLayoutEnabled = _maxChartsOtherWorkbooks > 1;
                            break;
                        case BatchExportObjects.ChartsAndShapes:
                            IsSingleItemsEnabled = _maxShapesOtherWorkbooks > 0;
                            IsSheetLayoutEnabled = _anyWorkbookMultipleChartsAndShapes;
                            break;
                        default:
                            throw new InvalidOperationException(String.Format(
                                "No case for {0} and {1}", Scope.AsString, Objects.AsString));
                    }
                    break;
                default:
                    throw new InvalidOperationException(
                        "No case defined for " + Scope.AsString);
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
                        "Cannot handle " + Objects.AsString);
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

        private int _numChartsActiveSheet;
        private int _numShapesActiveSheet;
        private int _maxChartsOtherSheets;
        private int _maxShapesOtherSheets;
        private int _maxChartsOtherWorkbooks;
        private int _maxShapesOtherWorkbooks;
        private bool _hasSheetWithMultipleChartsAndShapes;
        private bool _anyWorkbookMultipleChartsAndShapes;

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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Configuration;
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Excel.Instance;
using XLToolbox.WorkbookStorage;
using XLToolbox.Export.Models;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace XLToolbox.Export.ViewModels
{
    /// <summary>
    /// View model for the <see cref="Settings"/> class.
    /// </summary>
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    public class SingleExportSettingsViewModel : SettingsViewModelBase
    {
        #region Public properties

        /// <summary>
        /// Gets or sets the desired width of the exported graphic.
        /// </summary>
        public double Width
        {
            get { return ((SingleExportSettings)Settings).Width; }
            set
            {
                ((SingleExportSettings)Settings).Width = value;
                _dimensionsChanged = true;
                OnPropertyChanged("Width");
                if (PreserveAspect) OnPropertyChanged("Height");
            }
        }

        /// <summary>
        /// Gets or sets the desired width of the exported graphic.
        /// </summary>
        public double Height
        {
            get { return ((SingleExportSettings)Settings).Height; }
            set
            {
                ((SingleExportSettings)Settings).Height = value;
                _dimensionsChanged = true;
                OnPropertyChanged("Height");
                if (PreserveAspect) OnPropertyChanged("Width");
            }
        }

        /// <summary>
        /// Returns an enumerable list of available units and provides
        /// a bindable converter for a WPF ComboBox.
        /// </summary>
        public EnumProvider<Unit> Units
        {
            get
            {
                if (_unitString == null)
                {
                    _unitString = new EnumProvider<Unit>();
                    _unitString.PropertyChanged +=
                        (object sender, PropertyChangedEventArgs args) =>
                        {
                            if (args.PropertyName == "AsEnum")
                            {
                                ((SingleExportSettings)Settings).Unit = Units.AsEnum;
                                OnPropertyChanged("Width");
                                OnPropertyChanged("Height");
                            }
                            OnPropertyChanged("Units." + args.PropertyName);
                        };
                }
                return _unitString;
            }
        }

        /// <summary>
        /// Preserve aspect ratio if width or height are changed.
        /// </summary>
        public bool PreserveAspect
        {
            get { return ((SingleExportSettings)Settings).PreserveAspect; }
            set
            {
                ((SingleExportSettings)Settings).PreserveAspect = value;
                OnPropertyChanged("PreserveAspect");
            }
        }

        #endregion

        #region Commands

        /// <summary>
        /// Resets the Height and Width properties to the dimensions
        /// of the current selection in Excel.
        /// </summary>
        public DelegatingCommand ResetDimensionsCommand
        {
            get
            {
                if (_resetDimensionsCommand == null)
                {
                    _resetDimensionsCommand = new DelegatingCommand(
                        param => DoResetDimensions(),
                        param => CanResetDimensions()
                    );
                }
                return _resetDimensionsCommand;
            }
        }

        /// <summary>
        /// Causes the <see cref="ChooseFileNameMessage"/> to be sent.
        /// Upon confirmation of this message, the Export process will
        /// be started.
        /// </summary>
        public DelegatingCommand ChooseFileNameCommand
        {
            get
            {
                if (_chooseFileNameCommand == null)
                {
                    _chooseFileNameCommand = new DelegatingCommand(
                        param => DoChooseFileName(),
                        parma => CanChooseFileName());
                }
                return _chooseFileNameCommand;
            }
        }

        #endregion

        #region Messages

        public Message<FileNameMessageContent> ChooseFileNameMessage
        {
            get
            {
                if (_chooseFileNameMessage == null)
                {
                    _chooseFileNameMessage = new Message<FileNameMessageContent>();
                }
                return _chooseFileNameMessage;
            }
        }

        #endregion

        #region Constructors

        public SingleExportSettingsViewModel()
            : base()
        {
            if (PresetsRepository.Presets.Count == 0)
            {
                PresetsRepository.Presets.Add(new PresetViewModel());
            }
            if (ExcelInstance.Running)
            {
                SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
                if (ExcelInstance.Application.Workbooks.Count > 0)
                {
                    PresetsRepository.SelectLastUsedOrDefault(ExcelInstance.Application.ActiveWorkbook);
                }
                Settings = new SingleExportSettings(
                    PresetsRepository.SelectedPreset.RevealModelObject() as Preset,
                    svm.Bounds.Width, svm.Bounds.Height, true);
            }
            if (Settings == null)
            {
                Settings = new SingleExportSettings();
            }
            Units.AsEnum = Properties.Settings.Default.ExportUnit;
        }

        /*
        /// <summary>
        /// Instantiates the view model and adds the <paramref name="presetViewModel"/>
        /// to the Presets repository.
        /// </summary>
        /// <param name="presetViewModel">Preset view model to add to the repository.</param>
        public SingleExportSettingsViewModel(PresetViewModel presetViewModel)
            : this()
        {
            PresetsRepository.Presets.Add(presetViewModel);
        }

        public SingleExportSettingsViewModel(PresetViewModel preset, double width, double height)
            : this(preset)
        {
            Width = width;
            Height = height;
            _dimensionsChanged = false;
        }

        public SingleExportSettingsViewModel(PresetViewModel preset, double width, double height, bool preserveAspect)
            : this(preset, width, height)
        {
            PreserveAspect = preserveAspect;
        }

        public SingleExportSettingsViewModel(PresetViewModel preset, SelectionViewModel selection, bool preserveAspect)
            : this(preset)
        {
            Height = selection.Bounds.Height;
            Width = selection.Bounds.Width;
            _dimensionsChanged = false;
            PreserveAspect = preserveAspect;
        }
        */

        #endregion

        #region Implementation of SettingsViewModelBase

        /// <summary>
        /// Determins the suggested target directory and sends the
        /// ChooseFileNameMessage.
        /// </summary>
        protected override void DoExport()
        {
            if (CanExport())
            {
                // TODO: Make export asynchronous
                PresetsRepository.SaveSelected(ExcelInstance.Application.ActiveWorkbook);
                Properties.Settings.Default.ExportUnit = Units.AsEnum;
                Properties.Settings.Default.Save();
                Settings.Preset = SelectedPreset.RevealModelObject() as Preset;
                ProcessMessageContent pcm = new ProcessMessageContent();
                pcm.IsIndeterminate = true;
                ExportProcessMessage.Send(pcm);
                Exporter exporter = new Exporter();
                exporter.ExportSelection(Settings as SingleExportSettings);
                pcm.CompletedMessage.Send(pcm);
            }
        }

        protected override bool CanExport()
        {
            if (ExcelInstance.Running)
            {
                SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
                return (svm.Selection != null) && (SelectedPreset != null) &&
                    (Settings.Preset != null) && (Settings.Preset.Dpi > 0) &&
                    (Width > 0) && (Height > 0);
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region Private methods

        private void DoChooseFileName()
        {
            if (CanChooseFileName())
            {
                Preset preset = SelectedPreset.RevealModelObject() as Preset;
                ChooseFileNameMessage.Send(
                    new FileNameMessageContent(
                        GetExportPath(),
                        preset.FileType.ToFileFilter()
                        ),
                    (content) => DoConfirmFileName(content)
                );
            }
        }

        private bool CanChooseFileName()
        {
            return CanExport();
        }

        /// <summary>
        /// Called by Message.Respond() if the user has confirmed a file name
        /// in a view subscribed to the ChooseFileNameMessage. Triggers the
        /// actual export with the file name contained in the message content.
        /// </summary>
        /// <param name="messageContent"></param>
        private void DoConfirmFileName(FileNameMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                ((SingleExportSettings)Settings).FileName = messageContent.Value;
                DoExport();
            }
        }

        private void DoResetDimensions()
        {
            if (CanResetDimensions())
            {
                SelectionViewModel selection = new SelectionViewModel(
                    Excel.Instance.ExcelInstance.Application);
                bool oldAspectSwitch = PreserveAspect;
                PreserveAspect = false;
                Width = Unit.Point.ConvertTo(selection.Bounds.Width, Units.AsEnum);
                Height = Unit.Point.ConvertTo(selection.Bounds.Height, Units.AsEnum);
                PreserveAspect = oldAspectSwitch;
                _dimensionsChanged = false;
            }
        }

        private bool CanResetDimensions()
        {
            return _dimensionsChanged;
        }

        #endregion

        #region Private fields

        DelegatingCommand _chooseFileNameCommand;
        DelegatingCommand _resetDimensionsCommand;
        bool _dimensionsChanged;
        EnumProvider<Unit> _unitString;
        private Message<FileNameMessageContent> _chooseFileNameMessage;

        #endregion
    }
}

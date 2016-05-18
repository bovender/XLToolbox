/* SingleExportSettingsViewModel.cs
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
using System.ComponentModel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Export.Models;

namespace XLToolbox.Export.ViewModels
{
    /// <summary>
    /// View model for the <see cref="Settings"/> class.
    /// </summary>
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
            : this(new SingleExportSettings())
        { }

        public SingleExportSettingsViewModel(SingleExportSettings singleExportSettings)
            : base()
        {
            Settings = singleExportSettings;
            PresetViewModels.Select(Settings.Preset);
            // Need to explicitly set the selected enum value in the EnumProvider<Unit> collection.
            Units.AsEnum = singleExportSettings.Unit;
        }

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
                Logger.Info("DoExport");
                // TODO: Make export asynchronous
                SelectedPreset.Store();
                UserSettings.UserSettings.Default.ExportUnit = Units.AsEnum;
                SaveExportPath();
                Settings.Preset = SelectedPreset.RevealModelObject() as Preset;
                ProcessMessageContent pcm = new ProcessMessageContent();
                pcm.IsIndeterminate = true;
                Logger.Info("Send process message");
                ExportProcessMessage.Send(pcm);
                Exporter exporter = new Exporter();
                Logger.Info("Export selection");
                exporter.ExportSelection(Settings as SingleExportSettings);
                Logger.Info("Send completed message");
                pcm.CompletedMessage.Send(pcm);
            }
        }

        protected override bool CanExport()
        {
            SelectionViewModel svm = new SelectionViewModel(Instance.Default.Application);
            return (svm.Selection != null) && (SelectedPreset != null) &&
                (Settings.Preset != null) && (Settings.Preset.Dpi > 0) &&
                (Width > 0) && (Height > 0);
        }

        #endregion

        #region Overrides

        protected override void SaveExportPath()
        {
            base.SaveExportPath();
            UserSettings.UserSettings.Default.ExportPath =
                System.IO.Path.GetDirectoryName(FileName);
        }

        #endregion

        #region Private methods

        private void DoChooseFileName()
        {
            Logger.Info("DoChooseFileName");
            if (CanChooseFileName())
            {
                Preset preset = SelectedPreset.RevealModelObject() as Preset;
                ChooseFileNameMessage.Send(
                    new FileNameMessageContent(
                        LoadExportPath(),
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
            Logger.Info("DoConfirmFileName");
            if (messageContent.Confirmed)
            {
                Logger.Info("Confirmed");
                ((SingleExportSettings)Settings).FileName = messageContent.Value;
                UserSettings.UserSettings.Default.ExportPath =
                    System.IO.Path.GetDirectoryName(messageContent.Value);
                DoExport();
            }
            else
            {
                Logger.Info("Not confirmed");
            }
        }

        private void DoResetDimensions()
        {
            Logger.Info("DoResetDimensions");
            if (CanResetDimensions())
            {
                SelectionViewModel selection = new SelectionViewModel(
                    Instance.Default.Application);
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

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

/* SettingsViewModelBase.cs
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
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;
using XLToolbox.WorkbookStorage;
using XLToolbox.Export.Models;

namespace XLToolbox.Export.ViewModels
{
    /// <summary>
    /// Abstract base class for the <see cref="SingleExportSettingsViewModel"/>
    /// and <see cref="BatchExportSettingsViewModel"/> classes.
    /// </summary>
    /// <remarks>
    /// Rather than exposing a single PresetViewModel to subscribers, this
    /// view model has a PresetsRepository property so that subscribed views
    /// can select a preset from a preset repository. The last selected preset
    /// will be relayed to the wrapped Settings object.
    /// </remarks>
    public abstract class SettingsViewModelBase : ProcessViewModelBase
    {
        #region Public properties

        public PresetViewModel SelectedPreset
        {
            get
            {
                return PresetViewModels.SelectedViewModel;
            }
            set
            {
                PresetViewModels.SelectedViewModel = value;
                if (value != null)
                {
                    Settings.Preset = value.RevealModelObject() as Preset;
                }
                else
                {
                    Settings.Preset = null;
                }
                OnPropertyChanged("SelectedPreset");
            }
        }

        public PresetsRepositoryViewModel PresetViewModels
        {
            get
            {
                if (_presetsRepositoryViewModel == null)
                {
                    _presetsRepositoryViewModel = new PresetsRepositoryViewModel();
                    _presetsRepositoryViewModel.PropertyChanged += PresetViewModels_PropertyChanged;
                }
                return _presetsRepositoryViewModel;
            }
        }

        public string FileName
        {
            get { return Settings.FileName; }
            set
            {
                Settings.FileName = value;
                OnPropertyChanged("FileName");
            }
        }
        
        #endregion

        #region Protected properties 

        protected Settings Settings { get; set; }

        protected Exporter Exporter
        {
            get
            {
                if (_exporter == null)
                {
                    _exporter = new Exporter();
                    _exporter.ExportProgressCompleted += Exporter_ExportProgressCompleted;
                    _exporter.ExportFailed += Exporter_ExportFailed;
                }
                return _exporter;
            }
        }

        #endregion

        #region Commands

        public DelegatingCommand EditPresetsCommand
        {
            get
            {
                if (_editPresetsCommand == null)
                {
                    _editPresetsCommand = new DelegatingCommand(param => DoEditPresets());
                }
                return _editPresetsCommand;
            }
        }

        public DelegatingCommand ExportCommand
        {
            get
            {
                if (_exportCommand == null)
                {
                    _exportCommand = new DelegatingCommand(
                        param => DoExport(),
                        param => CanExport()
                    );
                }
                return _exportCommand;
            }
        }

        #endregion

        #region Messages

        public Message<ViewModelMessageContent> EditPresetsMessage
        {
            get
            {
                if (_editPresetsMessage == null)
                {
                    _editPresetsMessage = new Message<ViewModelMessageContent>();
                }
                return _editPresetsMessage;
            }
        }

        public Message<ProcessMessageContent> ExportProcessMessage
        {
            get
            {
                if (_exportProcessMessage == null)
                {
                    _exportProcessMessage = new Message<ProcessMessageContent>();
                }
                return _exportProcessMessage;
            }
        }

        #endregion

        #region Constructor

        public SettingsViewModelBase()
            : base()
        { }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Called when the ExportCommand is executed, triggers the
        /// export process.
        /// </summary>
        protected abstract void DoExport();

        protected abstract bool CanExport();

        #endregion

        #region Protected methods

        /// <summary>
        /// Returns the default path for export; this is either the path that was previously
        /// used to export from the current workbook, or the last used path, or the path of
        /// the current workbook, or the path for 'My Documents'.
        /// </summary>
        /// <returns>Default export path.</returns>
        protected string LoadExportPath()
        {
            Workbook wb = Excel.ViewModels.Instance.Default.ActiveWorkbook;
            Store store = new Store(wb);
            string defaultPath = UserSettings.UserSettings.Default.ExportPath;
            if (String.IsNullOrEmpty(defaultPath))
            {
                if (wb != null && !String.IsNullOrEmpty(wb.Path))
                {
                    defaultPath = System.IO.Path.GetDirectoryName(wb.Path);
                }
                else
                {
                    defaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
            }
            return store.Get(Properties.StoreNames.Default.ExportPath, defaultPath);
        }

        /// <summary>
        /// Saves the current export path for reuse.
        /// </summary>
        protected virtual void SaveExportPath()
        {
            Workbook wb = Excel.ViewModels.Instance.Default.ActiveWorkbook;
            using (Store store = new Store(wb))
            {
                store.Put(Properties.StoreNames.Default.ExportPath, Settings.FileName);
            }
        }

        protected virtual void DoEditPresets()
        {
            EditPresetsMessage.Send(
                new ViewModelMessageContent(PresetViewModels),
                content => OnPropertyChanged("Presets")
            );
        }

        #endregion

        #region Implementation of ViewModelBase and ProcessViewModelBase

        public override object RevealModelObject()
        {
            return Settings;
        }

        protected override bool IsProcessing()
        {
            return Exporter.IsProcessing;
        }

        protected override int GetPercentCompleted()
        {
            return Exporter.PercentCompleted;
        }

        protected override void CancelProcess()
        {
            Exporter.CancelExport();
        }

        #endregion

        #region Private methods

        private void PresetViewModels_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "SelectedViewModel")
            {
                OnPropertyChanged("SelectedPreset");
            }
        }

        private void Exporter_ExportProgressCompleted(object sender, EventArgs e)
        {
            Logger.Info("Export progress completed");
            ProcessMessageContent.CompletedMessage.Send();
            Logger.Info("... ProcessMessageContent.CompletedMessage was sent");
        }

        private void Exporter_ExportFailed(object sender, System.IO.ErrorEventArgs e)
        {
            Logger.Info("Exporter raised ExportFailed event, now sending my own message ...");
            ProcessFailedMessage.Send(
                new StringMessageContent(
                    String.Format(Strings.Export,
                    e.GetException().Message)));
            Logger.Info("... ExportFailedMessage was sent");
        }
        
        #endregion

        #region Private fields
        
        DelegatingCommand _exportCommand;
        DelegatingCommand _editPresetsCommand;
        Message<ViewModelMessageContent> _editPresetsMessage;
        Message<ProcessMessageContent> _exportProcessMessage;
        PresetsRepositoryViewModel _presetsRepositoryViewModel;
        Exporter _exporter;

        #endregion

        #region Class logger

        protected static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

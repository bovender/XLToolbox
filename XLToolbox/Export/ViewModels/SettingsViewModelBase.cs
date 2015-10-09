/* SettingsViewModelBase.cs
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
    public abstract class SettingsViewModelBase : ViewModelBase
    {
        #region Public properties

        public PresetViewModel SelectedPreset
        {
            get
            {
                return PresetsRepository.SelectedPreset;
            }
            set
            {
                PresetsRepository.SelectedPreset = value;
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

        public PresetsRepositoryViewModel PresetsRepository
        {
            get
            {
                if (_presetsRepositoryViewModel == null)
                {
                    _presetsRepositoryViewModel = new PresetsRepositoryViewModel();
                }
                return _presetsRepositoryViewModel;
            }
            set
            {
                _presetsRepositoryViewModel = value;
                OnPropertyChanged("PresetsRepository");
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
        {
            PresetsRepository = new PresetsRepositoryViewModel();
            PresetsRepository.PropertyChanged += PresetsRepository_PropertyChanged;
        }

        void PresetsRepository_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "SelectedPreset")
            {
                OnPropertyChanged("SelectedPreset");
            }
        }

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
            string defaultPath = Properties.Settings.Default.ExportPath;
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

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return Settings;
        }

        #endregion

        #region Private methods

        private void DoEditPresets()
        {
            EditPresetsMessage.Send(
                new ViewModelMessageContent(PresetsRepository),
                content => OnPropertyChanged("Presets")
            );
        }
        
        #endregion

        #region Private fields
        
        DelegatingCommand _exportCommand;
        DelegatingCommand _editPresetsCommand;
        Message<ViewModelMessageContent> _editPresetsMessage;
        Message<ProcessMessageContent> _exportProcessMessage;
        PresetsRepositoryViewModel _presetsRepositoryViewModel;

        #endregion
    }
}

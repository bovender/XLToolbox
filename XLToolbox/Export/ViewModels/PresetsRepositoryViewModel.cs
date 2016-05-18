/* PresetsRepositoryViewModel.cs
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
using System.Text;
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;
using System.Collections.Specialized;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Export.Models;
using XLToolbox.WorkbookStorage;

namespace XLToolbox.Export.ViewModels
{
    /// <summary>
    /// View model for an export settings repository.
    /// </summary>
    public class PresetsRepositoryViewModel : ViewModelBase
    {
        #region Public properties

        public PresetViewModelCollection ViewModels { get; private set; }

        public PresetViewModel SelectedViewModel
        {
            get
            {
                return ViewModels.LastSelected;
            }
            set
            {
                if (ViewModels.LastSelected != null)
                {
                    ViewModels.LastSelected.IsSelected = false;
                }
                if (value != null)
                {
                    value.IsSelected = true;
                }
                // No need to raise PropertyChanged here,
                // because we listen to the PresetViewModel's
                // event and relay it (in Presets_ViewModelPropertyChanged).
            }
        }

        #endregion

        #region Constructor

        public PresetsRepositoryViewModel()
            : base()
        {
            ViewModels = new PresetViewModelCollection();
            ViewModels.ViewModelPropertyChanged += Presets_ViewModelPropertyChanged;
        }

        #endregion

        #region Commands

        public DelegatingCommand AddCommand
        {
            get
            {
                if (_addCommand == null)
                {
                    _addCommand = new DelegatingCommand(
                        (param) => DoAddPreset());
                }
                return _addCommand;
            }
        }

        public DelegatingCommand RemoveCommand
        {
            get
            {
                if (_removeCommand == null)
                {
                    _removeCommand = new DelegatingCommand(
                        (param) => DoDeletePreset(),
                        (param) => CanDeletePreset());
                }
                return _removeCommand;
            }
        }

        public DelegatingCommand EditCommand
        {
            get
            {
                if (_editCommand == null)
                {
                    _editCommand = new DelegatingCommand(
                        (param) => DoEditPreset(),
                        (param) => CanEditPreset());
                }
                return _editCommand;
            }
        }

        #endregion

        #region Messages

        public Message<MessageContent> ConfirmRemoveMessage
        {
            get
            {
                if (_confirmRemoveMessage == null)
                {
                    _confirmRemoveMessage = new Message<MessageContent>();
                };
                return _confirmRemoveMessage;
            }
        }

        /// <summary>
        /// Sends a message indicating that a particular view model
        /// should be viewed for editing. The ExportSettingsViewModel object
        /// is conveyed in the message content.
        /// </summary>
        public Message<ViewModelMessageContent> EditSettingsMessage
        {
            get
            {
                if (_editMessage == null)
                {
                    _editMessage = new Message<ViewModelMessageContent>();
                };
                return _editMessage;
            }
        }

        #endregion

        #region Public methods

        public void SelectLastUsedOrDefault(Workbook workbook)
        {
            Preset preset = Preset.FromLastUsed(Excel.ViewModels.Instance.Default.ActiveWorkbook);
            if (preset == null)
            {
                preset = PresetsRepository.Default.First;
            }
            Select(preset);
        }

        public void Select(Preset preset)
        {
            if (preset == null)
            {
                throw new ArgumentNullException("preset", "Cannot select PresetViewModel without Preset");
            }
            PresetViewModel pvm = ViewModels.FirstOrDefault(p => p.IsViewModelOf(preset));
            pvm.IsSelected = true;
        }

        /*
        /// <summary>
        /// Stores the name of the last selected preset in the workbook's
        /// storage area and in the user settings.
        /// </summary>
        /// <param name="workbook">Workbook to store the preset in.</param>
        public void SaveSelected(Workbook workbook)
        {
            if (SelectedPreset != null)
            {
                Store store = new Store(workbook);
                store.Put(STORAGEKEY, SelectedPreset.Name);
                Properties.Settings.Default.ExportPresetName = SelectedPreset.Name;
                Properties.Settings.Default.Save();
            }
        }
         */

        #endregion

        #region Overrides of ViewModelBase

        protected override void DoCloseView()
        {
            base.DoCloseView();
        }
        #endregion

        #region Private methods

        private void DoAddPreset()
        {
            PresetViewModel pvm = new PresetViewModel();
            foreach (PresetViewModel p in ViewModels) { p.IsSelected = false; }
            ViewModels.Add(pvm);
            pvm.IsSelected = true;
            OnPropertyChanged("Presets");
            OnPropertyChanged("SelectedPresets");
        }

        private void DoDeletePreset()
        {
            ConfirmRemoveMessage.Send(
                new MessageContent(),
                content => ConfirmDeletePreset(content)
            );
        }

        private void ConfirmDeletePreset(MessageContent messageContent)
        {
            Dispatcher.Invoke(new System.Action(
                () =>
                {
                    if (CanDeletePreset() && messageContent.Confirmed)
                    {
                        ViewModels.RemoveSelected();
                        OnPropertyChanged("Presets");
                        OnPropertyChanged("SelectedPreset");
                    }
                })
            );
        }

        private bool CanDeletePreset()
        {
            return (SelectedViewModel != null);
        }

        private void DoEditPreset()
        {
            EditSettingsMessage.Send(
                new ViewModelMessageContent(SelectedViewModel),
                content => OnPropertyChanged("Presets")
            );
        }

        private bool CanEditPreset()
        {
            return (SelectedViewModel != null);
        }

        private void Presets_ViewModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case "Name": OnPropertyChanged("ViewModels"); break;
                case "IsSelected": OnPropertyChanged("SelectedViewModel"); break;
            }
        }

        #endregion

        #region Private fields

        DelegatingCommand _addCommand;
        DelegatingCommand _removeCommand;
        DelegatingCommand _editCommand;
        Message<MessageContent> _confirmRemoveMessage;
        Message<ViewModelMessageContent> _editMessage;

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return PresetsRepository.Default;
        }

        #endregion
    }
}

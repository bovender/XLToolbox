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

        public PresetViewModelCollection Presets { get; private set; }

        public PresetViewModel SelectedPreset
        {
            get
            {
                return Presets.LastSelected;
            }
            set
            {
                Presets.LastSelected.IsSelected = false;
                value.IsSelected = true;
                OnPropertyChanged("SelectedPreset");
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Instantiates the view model and creates a new PresetRepository
        /// instance which will load previously saved values in the background.
        /// </summary>
        public PresetsRepositoryViewModel()
            : base()
        {
            _repository = new PresetsRepository();
            Presets = new PresetViewModelCollection(_repository);
            Presets.ViewModelPropertyChanged += Presets_ViewModelPropertyChanged;
        }

        void Presets_ViewModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case "Name": OnPropertyChanged("Presets"); break;
                case "IsSelected": OnPropertyChanged("SelectedPreset"); break;
            }
        }

        /// <summary>
        /// Instantiates the view model by creating a new repository instance
        /// (which loads previously saved values, if they exist) and adding
        /// the <paramref name="presetViewModel"/> to the repository.
        /// </summary>
        /// <param name="presetViewModel">Preset view model (and associated model)
        /// to add to the repository.</param>
        public PresetsRepositoryViewModel(PresetViewModel presetViewModel)
            : this()
        {
            Presets.Add(presetViewModel);
        }

        public PresetsRepositoryViewModel(PresetsRepository repository)
            : base()
        {
            _repository = repository;
            Presets = new PresetViewModelCollection(_repository);
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

        /// <summary>
        /// Selects the last used preset as stored in the workbook's
        /// WorkbookStorage area or in the user settings, if any.
        /// Selects the first preset in the collection if no previously
        /// stored preset is found.
        /// </summary>
        /// <param name="workbook">Workbook whose stored settings to
        /// search for a previously used preset.</param>
        /// <exception cref="InvalidOperationException">If no presets
        /// exist in the collection.</exception>
        public void SelectLastUsedOrDefault(Workbook workbook)
        {
            if (Presets.Count == 0)
            {
                throw new InvalidOperationException(
                    "Cannot select a preset because there are no presets in the collection.");
            }

            Store workbookStore = new Store(workbook);
            string presetName = workbookStore.Get(
                STORAGEKEY,
                Properties.Settings.Default.ExportPresetName);
            PresetViewModel pvm = null;
            if (!String.IsNullOrEmpty(presetName))
            {
                pvm = Presets.FirstOrDefault(p => p.Name == presetName);
            }
            if (pvm != null)
            {
                pvm.IsSelected = true;
            }
            else
            {
                Presets[0].IsSelected = true;
            }
        }

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

        #endregion

        #region Private methods

        private void DoAddPreset()
        {
            Preset s = new Preset();
            PresetViewModel svm = new PresetViewModel(s);
            Presets.Add(svm);
            svm.IsSelected = true;
            OnPropertyChanged("Presets");
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
            if (CanDeletePreset() && messageContent.Confirmed)
            {
                this.Presets.RemoveSelected();
                OnPropertyChanged("Presets");
            }
        }

        private bool CanDeletePreset()
        {
            return (SelectedPreset != null);
        }

        private void DoEditPreset()
        {
            EditSettingsMessage.Send(
                new ViewModelMessageContent(SelectedPreset),
                content => OnPropertyChanged("Presets")
            );
        }

        private bool CanEditPreset()
        {
            return (SelectedPreset != null);
        }

        #endregion

        #region Private fields

        PresetsRepository _repository;
        DelegatingCommand _addCommand;
        DelegatingCommand _removeCommand;
        DelegatingCommand _editCommand;
        Message<MessageContent> _confirmRemoveMessage;
        Message<ViewModelMessageContent> _editMessage;

        #endregion

        #region Private contants

        private const string STORAGEKEY = "ExportPreset";

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _repository;
        }

        #endregion
    }
}

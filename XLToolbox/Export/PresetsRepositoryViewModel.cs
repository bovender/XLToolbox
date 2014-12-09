using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;

namespace XLToolbox.Export
{
    /// <summary>
    /// View model for an export settings repository.
    /// </summary>
    public class PresetsRepositoryViewModel : ViewModelBase
    {
        #region Public properties

        public PresetViewModelCollection ExportSettings { get; private set; }

        #endregion

        #region Constructor

        public PresetsRepositoryViewModel()
            : base()
        {
            _repository = new PresetsRepository();
            ExportSettings = new PresetViewModelCollection(_repository);
        }

        public PresetsRepositoryViewModel(PresetsRepository repository)
            : base()
        {
            _repository = repository;
            ExportSettings = new PresetViewModelCollection(_repository);
        }

        #endregion

        #region Commands

        public DelegatingCommand AddSettingsCommand
        {
            get
            {
                if (_addSettingsCommand == null)
                {
                    _addSettingsCommand = new DelegatingCommand(
                        (param) => DoAddSettings());
                }
                return _addSettingsCommand;
            }
        }

        public DelegatingCommand RemoveSettingsCommand
        {
            get
            {
                if (_removeSettingsCommand == null)
                {
                    _removeSettingsCommand = new DelegatingCommand(
                        (param) => DoDeleteSettings(),
                        (param) => CanDeleteSettings());
                }
                return _removeSettingsCommand;
            }
        }

        public DelegatingCommand EditSettingsCommand
        {
            get
            {
                if (_editSettingsCommand == null)
                {
                    _editSettingsCommand = new DelegatingCommand(
                        (param) => DoEditSettings(),
                        (param) => CanEditSettings());
                }
                return _editSettingsCommand;
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
                if (_editSettingsMessage == null)
                {
                    _editSettingsMessage = new Message<ViewModelMessageContent>();
                };
                return _editSettingsMessage;
            }
        }

        #endregion

        #region Private methods

        private void DoAddSettings()
        {
            Export.Preset s = new Export.Preset();
            PresetViewModel svm = new PresetViewModel(s);
            ExportSettings.Add(svm);
            svm.IsSelected = true;
            OnPropertyChanged("ExportSettings");
        }

        private void DoDeleteSettings()
        {
            ConfirmRemoveMessage.Send(
                new MessageContent(),
                content => ConfirmDeleteSettings(content)
            );
        }

        private void ConfirmDeleteSettings(MessageContent messageContent)
        {
            if (CanDeleteSettings() && messageContent.Confirmed)
            {
                this.ExportSettings.RemoveSelected();
                OnPropertyChanged("ExportSettings");
            }
        }

        private bool CanDeleteSettings()
        {
            return (this.ExportSettings.CountSelected > 0);
        }

        private void DoEditSettings()
        {
            EditSettingsMessage.Send(
                new ViewModelMessageContent(ExportSettings.LastSelected),
                content => OnPropertyChanged("ExportSettings")
            );
        }

        private bool CanEditSettings()
        {
            return (this.ExportSettings.CountSelected > 0);
        }

        #endregion

        #region Private fields

        PresetsRepository _repository;
        DelegatingCommand _addSettingsCommand;
        DelegatingCommand _removeSettingsCommand;
        DelegatingCommand _editSettingsCommand;
        Message<MessageContent> _confirmRemoveMessage;
        Message<ViewModelMessageContent> _editSettingsMessage;

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _repository;
        }

        #endregion
    }
}

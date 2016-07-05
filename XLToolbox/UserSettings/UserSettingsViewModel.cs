/* UserSettingsViewModel.cs
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
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;

namespace XLToolbox.UserSettings
{
    public class UserSettingsViewModel : ViewModelBase
    {
        #region Properties

        public int TaskPaneWidth
        {
            get
            {
                return _taskPaneWidth;
            }
            set
            {
                _taskPaneWidth = value;
                OnPropertyChanged("TaskPaneWidth");
            }
        }

        public int MaxTaskPaneWidth
        {
            get
            {
                return 600;
            }
        }

        public int MinTaskPaneWidth
        {
            get
            {
                return 200;
            }
        }

        public bool IsLoggingEnabled
        {
            get
            {
                return _isLoggingEnabled;
            }
            set
            {
                _isLoggingEnabled = value;
                OnPropertyChanged("IsLoggingEnabled");
            }
        }

        public string ProfileFolderPath
        {
            get
            {
                return System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    Properties.Settings.Default.AppDataFolder,
                    Properties.Settings.Default.UserFolder);
            }
        }

        public EnumProvider<Language> Language
        {
            get
            {
                if (_languageProvider == null)
                {
                    _languageProvider = new Bovender.Mvvm.EnumProvider<Language>(_language);
                    _languageProvider.PropertyChanged += (sender, args) =>
                    {
                        _language = _languageProvider.AsEnum;
                        _dirty |= _language != _originalLanguage;
                    };
                }
                return _languageProvider;
            }
        }

        #endregion

        #region Commands

        public DelegatingCommand SaveCommand
        {
            get
            {
                if (_saveCommand == null)
                {
                    _saveCommand = new DelegatingCommand(
                        param => DoSave(),
                        param => CanSave());
                }
                return _saveCommand;
            }
        }

        public DelegatingCommand OpenProfileFolderCommand
        {
            get
            {
                if (_openProfileFolderCommand == null)
                {
                    _openProfileFolderCommand = new DelegatingCommand(
                        param => DoOpenProfileFolder());
                }
                return _openProfileFolderCommand;
            }
        }

        public DelegatingCommand EditLegacyPreferencesCommand
        {
            get
            {
                if (_editLegacyPreferences == null)
                {
                    _editLegacyPreferences = new DelegatingCommand(
                        param => DoOpenLegacyPreferences());
                }
                return _editLegacyPreferences;
            }
        }

        #endregion

        #region Messages

        public Message<MessageContent> RestartToTakeEffectMessage
        {
            get
            {
                if (_restartToChange == null)
                {
                    _restartToChange = new Message<MessageContent>();
                }
                return _restartToChange;
            }
        }

        #endregion

        #region Constructor

        public UserSettingsViewModel()
        {
            UserSettings u = UserSettings.Default;
            _isLoggingEnabled = u.EnableLogging;
            if (!Enum.TryParse(u.LanguageCode, true, out _language))
            {
                Logger.Warn("UserSettingsViewModel: Could not parse language code to enum, falling back to default");
                _language = XLToolbox.UserSettings.Language.En;
            }
            _originalLanguage = _language;
            if (XLToolbox.SheetManager.TaskPaneManager.InitializedAndVisible)
            {
                _taskPaneWidth = XLToolbox.SheetManager.TaskPaneManager.Default.Width;
            }
            else
            {
                _taskPaneWidth = u.TaskPaneWidth;
            }
            PropertyChanged += (sender, args) =>
            {
                _dirty = true;
            };
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return UserSettings.Default;
        }

        #endregion

        #region Private methods

        private void DoSave()
        {
            Logger.Info("DoSave");
            _dirty = false;
            UserSettings u = UserSettings.Default;
            u.TaskPaneWidth = TaskPaneWidth;
            u.EnableLogging = IsLoggingEnabled;
            u.LanguageCode = Language.SelectedItem.Value.ToString();
            if (XLToolbox.SheetManager.TaskPaneManager.InitializedAndVisible)
            {
                XLToolbox.SheetManager.TaskPaneManager.Default.Width = _taskPaneWidth;
            }
            if (_originalLanguage != _language)
            {
                OnRestartToTakeEffect();
            }
            DoCloseView();
        }

        private bool CanSave()
        {
            return _dirty;
        }

        private void DoOpenLegacyPreferences()
        {
            Logger.Info("Open legacy preferences");
            if (!_dirty)
            {
                Logger.Info("(Not dirty, closing view.)");
                DoCloseView();
            }
            XLToolbox.Dispatcher.Execute(Command.LegacyPrefs);
        }

        private void DoOpenProfileFolder()
        {
            Logger.Info("Open profile folder");
            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(ProfileFolderPath));
        }

        protected void OnRestartToTakeEffect()
        {
            RestartToTakeEffectMessage.Send();
        }

        #endregion

        #region Private fields

        private DelegatingCommand _saveCommand;
        private DelegatingCommand _openProfileFolderCommand;
        private DelegatingCommand _editLegacyPreferences;
        private Message<MessageContent> _restartToChange;
        private bool _dirty;
        private int _taskPaneWidth;
        private bool _isLoggingEnabled;
        private Language _originalLanguage;
        private Language _language;
        private EnumProvider<Language> _languageProvider;
        
        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

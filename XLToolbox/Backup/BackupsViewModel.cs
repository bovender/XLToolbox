/* BackupsViewModel.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Backup
{
    public class BackupsViewModel : ViewModelBase
    {
        #region Properties

        public bool IsEnabled
        {
            get
            {
                return Backups.IsEnabled;
            }
            set
            {
                Backups.IsEnabled = value;
                OnPropertyChanged("IsEnabled");
                OnPropertyChanged("FlashBackupsDisclaimer");
            }
        }

        public bool FlashBackupsDisclaimer
        {
            get
            {
                return IsEnabled && !_wasEnabled;
            }
        }

        public bool HasBackups
        {
            get
            {
                return (BackupFiles != null) && (BackupFiles.Count > 0);
            }
        }

        public BackupFilesCollection BackupFiles { get; private set; }

        public string BackupDir { get; private set; }

        #endregion

        #region MVVM commands

        public DelegatingCommand OpenBackupCommand
        {
            get
            {
                if (_openBackupCommand == null)
                {
                    _openBackupCommand = new DelegatingCommand(DoOpenBackup, CanOpenBackup);
                }
                return _openBackupCommand;
            }
        }

        public DelegatingCommand OpenBackupDirCommand
        {
            get
            {
                if (_openBackupDirCommand == null)
                {
                    _openBackupDirCommand = new DelegatingCommand(DoOpenBackupDir);
                }
                return _openBackupDirCommand;
            }
        }

        public DelegatingCommand DeleteBackupCommand
        {
            get
            {
                if (_deleteBackupCommand == null)
                {
                    _deleteBackupCommand = new DelegatingCommand(DoDeleteBackup, CanDeleteBackup);
                }
                return _deleteBackupCommand;
            }
        }

        public DelegatingCommand DeleteAllBackupsCommand
        {
            get
            {
                if (_deleteAllBackupsCommand == null)
                {
                    _deleteAllBackupsCommand = new DelegatingCommand(ConfirmDeleteAllBackups, CanDeleteAllBackups);
                }
                return _deleteAllBackupsCommand;
            }
        }

        #endregion

        #region MVVM messages

        public Message<ViewModelMessageContent> ConfirmDeleteAllBackupsMessage
        {
            get
            {
                if (_confirmDeleteAllBackupsMessage == null)
                {
                    _confirmDeleteAllBackupsMessage = new Message<ViewModelMessageContent>();
                }
                return _confirmDeleteAllBackupsMessage;
            }
        }

        #endregion

        #region Constructors

        public BackupsViewModel(Workbook workbook)
        {
            _wasEnabled = IsEnabled;
            string dir = UserSettings.UserSettings.Default.BackupDir;
            BackupDir = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(workbook.FullName),
                dir);
            _backups = new Backups(workbook.FullName, dir);
            if (_backups.Files != null)
            {
                BackupFiles = new BackupFilesCollection(_backups);
            }
        }

        #endregion

        #region Private methods

        private void DoOpenBackup(object param)
        {
            if (CanOpenBackup(null))
            {
                Logger.Info("DoOpenBackup");
                BackupFiles.LastSelected.OpenCommand.Execute(null);
                CloseViewCommand.Execute(null);
            }
        }

        private bool CanOpenBackup(object param)
        {
            return (BackupFiles != null) && (BackupFiles.LastSelected != null);
        }

        private void DoOpenBackupDir(object param)
        {
            Logger.Info("DoOpenBackupDir");
            System.Diagnostics.Process.Start(BackupDir);
        }

        private void DoDeleteBackup(object param)
        {
            if (CanDeleteBackup(null))
            {
                BackupFileViewModel bf = BackupFiles.LastSelected;
                Logger.Info("DoDeleteBackup: Executing BackupFile.DeleteCommand");
                bf.DeleteCommand.Execute(null);
                if (bf.IsDeleted)
                {
                    Logger.Info("DoDeleteBackup: Removing item from list");
                    BackupFiles.Remove(BackupFiles.LastSelected);
                    OnPropertyChanged("BackupFiles");
                }
                else
                {
                    Logger.Info("DoDeleteBackup: BackupFile is not deleted, keeping in list");
                }
            }
        }

        private bool CanDeleteBackup(object param)
        {
            return (BackupFiles != null) && (BackupFiles.LastSelected != null);
        }

        private void ConfirmDeleteAllBackups(object param)
        {
            ConfirmDeleteAllBackupsMessage.Send(new ViewModelMessageContent(this), DoDeleteAllBackups);
        }

        private bool CanDeleteAllBackups(object param)
        {
            return (BackupFiles != null) && (BackupFiles.Count > 0);
        }

        private void DoDeleteAllBackups(ViewModelMessageContent messageContent)
        {
            if (messageContent.Confirmed)
            {
                Logger.Info("DoDeleteAllBackups: Iterating over {0} BackupFiles", BackupFiles.Count);
                foreach (BackupFileViewModel vm in BackupFiles)
                {
                    vm.DeleteCommand.Execute(null);
                    vm.IsSelected = vm.IsDeleted;
                }
                Logger.Info("DoDeleteAllBackups: Purging list");
                BackupFiles.RemoveSelected();
                Logger.Info("DoDeleteAllBackups: {0} BackupFiles remaining", BackupFiles.Count);
                OnPropertyChanged("HasBackups");
            }
            else
            {
                Logger.Info("DoDeleteAllBackups: Action was not confirmed, not deleting backups");
            }
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return _backups;
        }

        #endregion

        #region Fields

        private bool _wasEnabled;
        private Backups _backups;
        private DelegatingCommand _openBackupCommand;
        private DelegatingCommand _deleteBackupCommand;
        private DelegatingCommand _deleteAllBackupsCommand;
        private DelegatingCommand _openBackupDirCommand;
        private Message<ViewModelMessageContent> _confirmDeleteAllBackupsMessage;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

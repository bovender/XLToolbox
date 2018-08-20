/* BackupFileViewModel.cs
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
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm;

namespace XLToolbox.Backup
{
    public class BackupFileViewModel : ViewModelBase
    {
        #region Properties
		
        public BackupFile BackupFile { get; private set; }

        public bool IsDeleted { get { return BackupFile.IsDeleted; } }

	    #endregion

        #region MVVM commands

        public DelegatingCommand OpenCommand
        {
            get
            {
                if (_openCommand == null)
                {
                    _openCommand = new DelegatingCommand(DoOpen);
                }
                return _openCommand;
            }
        }

        public DelegatingCommand DeleteCommand
        {
            get
            {
                if (_deleteCommand == null)
                {
                    _deleteCommand = new DelegatingCommand(DoDelete);
                }
                return _deleteCommand;
            }
        }

        #endregion

        #region Constructors

        public BackupFileViewModel() : base() { }

        public BackupFileViewModel(BackupFile backupFile)
            : this()
        {
            BackupFile = backupFile;
        }

        public BackupFileViewModel(string path)
            : this(new BackupFile(path))
        { }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return BackupFile;
        }

        public override string DisplayString
        {
            get
            {
                if (BackupFile != null && BackupFile.TimeStamp != null)
                {
                    string datePart;
                    int dayDiff = DateTime.Today.Subtract(BackupFile.TimeStamp.DateTime.Date).Days;
                    switch (dayDiff)
                    {
                        case 0:
                            datePart = Strings.Today;
                            break;
                        case 1:
                            datePart = Strings.Yesterday;
                            break;
                        case 2:
                        case 3:
                        case 4:
                        case 5:
                        case 6:
                        case 7:
                            datePart = BackupFile.TimeStamp.DateTime.ToString("dddd");
                            break;
                        default:
                            datePart = BackupFile.TimeStamp.DateTime.ToShortDateString();
                            break;
                    }
                    return String.Format("{0} {1}", datePart, BackupFile.TimeStamp.DateTime.ToShortTimeString());
                }
                else
                {
                    return "n/a";
                }
            }
        }

        #endregion

        #region Private methods

        private void DoOpen(object param)
        {
            BackupFile.Open();
        }

        private void DoDelete(object param)
        {
            BackupFile.Delete();
        }

        #endregion

        #region Fields

        DelegatingCommand _openCommand;
        DelegatingCommand _deleteCommand;

        #endregion
    }
}

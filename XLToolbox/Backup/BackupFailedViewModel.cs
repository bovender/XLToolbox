/* BackupFailedViewModel.cs
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
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Backup
{
    public class BackupFailedViewModel : ViewModelBase
    {
        #region Properties

        public Exception Exception { get; private set; }

        public string Message
        {
            get
            {
                if (Exception != null)
                {
                    return Exception.Message;
                }
                else
                {
                    return String.Empty;
                }
            }
        }

        public bool Suppress
        {
            get
            {
                return XLToolbox.UserSettings.UserSettings.Default.SuppressBackupFailureMessage;
            }
            set
            {
                XLToolbox.UserSettings.UserSettings.Default.SuppressBackupFailureMessage = value;
                OnPropertyChanged("Suppress");
            }
        }

        #endregion

        #region Constructors

        public BackupFailedViewModel(Exception exception)
        {
            Exception = exception;
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return Exception;
        }

        #endregion
    }
}

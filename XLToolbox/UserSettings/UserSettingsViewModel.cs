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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.UserSettings
{
    public class UserSettingsViewModel : ViewModelBase
    {
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

        #endregion

        #region Messages
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
            _dirty = false;
        }

        private bool CanSave()
        {
            return _dirty;
        }

        private void DoOpenProfileFolder()
        {
            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(
                    System.IO.Path.GetDirectoryName(
                        UserSettings.Default.GetSettingsFilePath())));
        }

        #endregion

        #region Private fields

        private DelegatingCommand _saveCommand;
        private DelegatingCommand _openProfileFolderCommand;
        private bool _dirty;
        
        #endregion
    }
}

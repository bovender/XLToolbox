/* LogFileViewModel.cs
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
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm;

namespace XLToolbox.Logging
{
    public class LogFileViewModel : ViewModelBase
    {
        #region Properties

        public string CurrentLog
        {
            get
            {
                return LogFile.Default.CurrentLog;
            }
        }

        #endregion

        #region Commands

        public DelegatingCommand EnableLoggingCommand
        {
            get
            {
                if (_enableLoggingCommand == null)
                {
                    _enableLoggingCommand = new DelegatingCommand(
                        param => DoEnableLogging());
                }
                return _enableLoggingCommand;
            }
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return LogFile.Default;
        }

        #endregion

        #region Private methods

        private void DoEnableLogging()
        {
            LogFile.Default.IsFileLoggingEnabled = true;
            Logger.Info("User enabled logging after incomplete shutdown");
            DoCloseView();
        }

        #endregion

        #region Private fields

        DelegatingCommand _enableLoggingCommand;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

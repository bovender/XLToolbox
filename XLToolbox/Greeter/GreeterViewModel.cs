/* GreeterViewModel.cs
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
using Bovender.Versioning;

namespace XLToolbox.Greeter
{
    /// <summary>
    /// View model for the greeter screen.
    /// </summary>
    public class GreeterViewModel : About.AboutViewModel
    {
        #region Commands

        public DelegatingCommand WhatsNewCommand
        {
            get
            {
                if (_whatsNewCommand == null)
                {
                    _whatsNewCommand = new DelegatingCommand(
                        (param) => DoShowWhatsNew()
                        );
                }
                return _whatsNewCommand;
            }
        }

        public DelegatingCommand DonateCommand
        {
            get
            {
                if (_donateCommand == null)
                {
                    _donateCommand = new DelegatingCommand(
                        (param) => DoShowDonatePage()
                        );
                }
                return _donateCommand;
            }
        }

        #endregion

        #region Private methods

        private void DoShowWhatsNew()
        {
            System.Diagnostics.Process.Start(Properties.Settings.Default.WhatsNewUrl);
            CloseViewCommand.Execute(null);
        }

        private void DoShowDonatePage()
        {
            System.Diagnostics.Process.Start(Properties.Settings.Default.DonateUrl);
            CloseViewCommand.Execute(null);
        }

        #endregion

        #region Private fields

        private DelegatingCommand _whatsNewCommand;
        private DelegatingCommand _donateCommand;

        /*
        private Message<MessageContent> _whatsNewMessage;
        private Message<MessageContent> _donateMessage;
         */

        #endregion

        #region Overrides

        public override object RevealModelObject()
        {
            return null;
        }

        #endregion
    }
}

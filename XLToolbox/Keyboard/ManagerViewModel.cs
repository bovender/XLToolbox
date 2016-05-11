/* ManagerViewModel.cs
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

namespace XLToolbox.Keyboard
{
    public class ManagerViewModel : ViewModelBase
    {
        #region Properties

        public ShortcutViewModelCollection ShortcutViewModels
        {
            get
            {
                return _shortcutViewModels;
            }
        }

        #endregion

        #region Commands

        public DelegatingCommand EditShortcutCommand
        {
            get
            {
                if (_editShortcutCommand == null)
                {
                    _editShortcutCommand = new DelegatingCommand(
                        param => DoEditShortcut());
                }
                return _editShortcutCommand;
            }
        }

        #endregion

        #region Messages

        public Message<ViewModelMessageContent> EditShortcutMessage
        {
            get
            {
                if (_editShortcutMessage == null)
                {
                    _editShortcutMessage = new Message<ViewModelMessageContent>();
                }
                return _editShortcutMessage;
            }
        }

        #endregion

        #region Constructor

        public ManagerViewModel()
        {
            _shortcutViewModels = new ShortcutViewModelCollection();
        }

        #endregion

        #region Private methods

        private void DoEditShortcut()
        {
            EditShortcutMessage.Send(new ViewModelMessageContent(ShortcutViewModels.LastSelected));
        }

        #endregion

        #region Overrides

        public override object RevealModelObject()
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Private fields

        ShortcutViewModelCollection _shortcutViewModels;
        DelegatingCommand _editShortcutCommand;
        Message<ViewModelMessageContent> _editShortcutMessage;

        #endregion
    }
}

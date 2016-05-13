/* ShortcutViewModel.cs
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
using System.ComponentModel;

namespace XLToolbox.Keyboard
{
    public class ShortcutViewModel : ViewModelBase, IDataErrorInfo
    {
        #region Properties

        public string Command
        {
            get
            {
                if (_editedShortcut != null)
	            {
                    return _editedShortcut.Command.ToString();
	            }
                else
	            {
                    return String.Empty;
	            }
            }
        }

        public string HumanKeySequence
        {
            get
            {
                if (_editedShortcut != null)
                {
                    return _editedShortcut.HumanKeySequence;
                }
                else
                {
                    return String.Empty;
                }
            }
        }

        public string KeySequence
        {
            get
            {
                return _editedShortcut.KeySequence;
            }
            set
            {
                _editedShortcut.KeySequence = value;
                IsDirty = _editedShortcut.KeySequence != _shortcut.KeySequence;
                OnPropertyChanged("KeySequence");
                OnPropertyChanged("HumanKeySequence");
                OnPropertyChanged("IsValid");
            }
        }

        public bool IsValid
        {
            get
            {
                return _editedShortcut.IsValid;
            }
        }

        public bool IsDirty
        {
            get
            {
                return _dirty;
            }
            protected set
            {
                _dirty = value;
                OnPropertyChanged("IsDirty");
            }
        }

        #endregion

        #region Commmand

        public DelegatingCommand SaveShortcutCommand
        {
            get
            {
                if (_saveShortcutCommand == null)
                {
                    _saveShortcutCommand = new DelegatingCommand(
                        param => DoSaveShortcut(),
                        parma => CanSaveShortcut());
                }
                return _saveShortcutCommand;
            }
        }

        public DelegatingCommand ResetShortcutCommand
        {
            get
            {
                if (_resetShortcutCommand == null)
                {
                    _resetShortcutCommand = new DelegatingCommand(
                        param => DoResetShortcut(),
                        parma => CanResetShortcut());
                }
                return _resetShortcutCommand;
            }
        }

        #endregion

        #region Constructor

        public ShortcutViewModel(Shortcut shortcut)
        {
            _shortcut = shortcut;
            _editedShortcut = new Shortcut(_shortcut.KeySequence, _shortcut.Command);
        }
        
        #endregion

        #region Private methods

        private void DoSaveShortcut()
        {
            _shortcut.KeySequence = _editedShortcut.KeySequence;
            _dirty = false;
            CloseViewCommand.Execute(null);
        }

        private void DoResetShortcut()
        {
            KeySequence = _shortcut.KeySequence;
            _dirty = false;
        }

        private bool CanSaveShortcut()
        {
            return IsDirty && IsValid;
        }

        private bool CanResetShortcut()
        {
            return IsDirty;
        }

        #endregion

        #region Overrides

        public override object RevealModelObject()
        {
            return _shortcut;
        }

        public override string ToString()
        {
            return String.Format("{0} ({1})", _shortcut.Command.ToString(), _shortcut.HumanKeySequence);
        }

        #endregion

        #region Fields

        Shortcut _shortcut;
        Shortcut _editedShortcut;
        bool _dirty;
        DelegatingCommand _saveShortcutCommand;
        DelegatingCommand _resetShortcutCommand;

        #endregion

        string IDataErrorInfo.Error { get { throw new NotImplementedException(); } }

        string IDataErrorInfo.this[string columnName]
        {
            get
            {
                if (!IsValid)
                {
                    return "Invalid key sequence";
                }
                else
                {
                    return null;
                }
            }
        }
    }
}

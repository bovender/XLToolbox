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

namespace XLToolbox.Keyboard
{
    public class ShortcutViewModel : ViewModelBase
    {
        #region Properties

        public string Command
        {
            get
            {
                if (_shortcut != null)
	            {
                    return _shortcut.Command.ToString();
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
                if (_shortcut != null)
                {
                    return _shortcut.HumanKeySequence;
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
                return _shortcut.KeySequence;
            }
            set
            {
                _shortcut.KeySequence = value;
                OnPropertyChanged("KeySequence");
            }
        }

        #endregion

        #region Constructor

        public ShortcutViewModel(Shortcut shortcut)
        {
            _shortcut = shortcut;
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

        #endregion
    }
}

/* EnumViewModel.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
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

namespace Bovender.Mvvm.ViewModels
{
    /// <summary>
    /// View model for enum members. This class is used internally by the
    /// <see cref="EnumProvider"/>. For specific purposes, the a new class
    /// based on EnumProvider should be created. The EnumViewModel class
    /// is sealed and cannot be inherited from.
    /// </summary>
    public sealed class EnumViewModel<T> : ViewModelBase
        where T : struct, IConvertible
    {
        #region Public properties

        public T Value { get; private set; }
        public string Description
        {
            get
            {
                return _description;
            }
            set
            {
                _description = value;
                OnPropertyChanged("Description");
            }
        }

        public string ToolTip
        {
            get
            {
                return _toolTip;
            }
            set
            {
                _toolTip = value;
                OnPropertyChanged("ToolTip");
            }
        }

        public bool IsEnabled
        {
            get
            {
                return _isEnabled;
            }
            set
            {
                _isEnabled = value;
                OnPropertyChanged("IsEnabled");
            }
        }

        #endregion

        #region Constructors

        internal EnumViewModel(T value)
            : base()
        {
            Value = value;
            _isEnabled = true;
        }

        internal EnumViewModel(T value, string description)
            : this(value)
        {
            _description = description;
        }

        internal EnumViewModel(T value, string description, string toolTip)
            : this(value, description)
        {
            _toolTip = toolTip;
        }

        internal EnumViewModel(T value, string description, string toolTip,
            bool isEnabled)
            : this(value, description, toolTip)
        {
            _isEnabled = isEnabled;
        }

        internal EnumViewModel(T value, string description, bool isEnabled)
            : this(value, description)
        {
            _isEnabled = isEnabled;
        }

        #endregion

        #region Overrides

        public override string DisplayString
        {
            get
            {
                if (String.IsNullOrEmpty(Description))
                {
                    return Value.ToString();
                }
                else
                {
                    return Description;
                }
            }
        }

        public override string ToString()
        {
            return DisplayString;
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return Value;
        }
        
        #endregion

        #region Private fields

        private string _description;
        private string _toolTip;
        private bool _isEnabled;

        #endregion
    }
}

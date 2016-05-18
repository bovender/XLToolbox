/* Shortcut.cs
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
using System.Text.RegularExpressions;

namespace XLToolbox.Keyboard
{
    /// <summary>
    /// Keyboard shortcut for a XL Toolbox command.
    /// </summary>
    [Serializable]
    public class Shortcut
    {
        #region Public properties

        /// <summary>
        /// Gets or sets the XLToolbox.Command associated with the KeySequence.
        /// </summary>
        public Command Command { get; set; }

        /// <summary>
        /// Gets or sets the technical representation of a key sequence in Excel.
        /// </summary>
        /// <remarks>
        /// Modifiers:
        /// <list type="bullet">
        /// <item><term>^</term><description>CONTROL</description></item>
        /// <item><term>+</term><description>SHIFT</description></item>
        /// <item><term>%</term><description>ALT</description></item>
        /// <item></item>
        /// </list>
        /// </remarks>
        public string KeySequence
        {
            get
            {
                return _keySequence;
            }
            set
            {
                if (value != null)
                {
                    _keySequence = value.ToUpper();
                }
                else
                {
                    _keySequence = null;
                }
                Validate();
            }
        }

        /// <summary>
        /// Gets a human-readable representation of the key sequence.
        /// </summary>
        [YamlDotNet.Serialization.YamlIgnore]
        public string HumanKeySequence
        {
            get
            {
                if (IsValid)
                {
                    string s = Regex.Replace(KeySequence, @"(?<!{)\^(?!})", "CONTROL ");
                    s = Regex.Replace(s, @"(?<!{)\+(?!})", "SHIFT ");
                    s = Regex.Replace(s, @"(?<!{)\%(?!})", "ALT ");
                    return s.Replace("{+}", "+").Replace("{^}", "^").Replace("{%}", "%");
                }
                else
                {
                    return Strings.InvalidKeySequence;
                }
            }
        }

        [YamlDotNet.Serialization.YamlIgnore]
        public bool IsValid
        {
            get
            {
                // The match relies on the KeySequence having been converted ToUpper.
                return _isValid;
            }
        }
            
        #endregion

        #region Constructors

        public Shortcut() { }

        public Shortcut(string keySequence, Command command)
        {
            KeySequence = keySequence;
            Command = command;
        }

        #endregion

        #region Public methods

        public void Register()
        {
            // The VBA method `XltbCmd` is declared in XLToolboxKeyboardBridge.xlam.
            // Note the special quoting of commands that is required for OnKey to work.
            if (!string.IsNullOrEmpty(KeySequence) && IsValid)
            {
                Excel.ViewModels.Instance.Default.Application.OnKey(KeySequence, "'XltbCmd(\"" + Command + "\")'");
            }
        }

        public void Unregister()
        {
            if (!string.IsNullOrEmpty(KeySequence) && IsValid)
            {
                Excel.ViewModels.Instance.Default.Application.OnKey(KeySequence);
            }
        }

        #endregion

        #region Overrides

        public override string ToString()
        {
            return Command.ToString() + "(" + HumanKeySequence + ")";
        }

        #endregion

        #region Protected methods

        // protected void Execute()
        // {
        //     Dispatcher.Execute(Command);
        // }

        protected void Validate()
        {
            _isValid = String.IsNullOrEmpty(KeySequence) || _validationPattern.IsMatch(KeySequence);
        }

        #endregion

        #region Private fields

        private string _keySequence;
        private bool _isValid;

        #endregion

        #region Static fields

        private static readonly Regex _validationPattern = new Regex(
            @"^(\+?\^?%?|\+?%?\^?|%?\+?\^?|%?\^?\+?|\^?%?\+?|\^?\+?%?)" +
            @"([^+^%]|{(\+|\^|%|F[1-9]|F1[0-5]|ENTER|PGUP|PGDN|BS|BACKSPACE|BREAK|CAPSLOCK|CLEAR|DEL|DELETE" +
            @"|DOWN|END|~|ESC|ESCAPE|HELP|HOME|INSERT|LEFT|NUMLOCK|RIGHT|SCROLLLOCK|TAB|UP)})$");

        #endregion
    }
}

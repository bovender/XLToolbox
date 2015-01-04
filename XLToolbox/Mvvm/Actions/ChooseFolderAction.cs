/* ChooseFolderAction.cs
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
using System.Windows;
using Bovender.Mvvm.Actions;
using Bovender.Mvvm.Messaging;
using System.Windows.Forms;

namespace XLToolbox.Mvvm.Actions
{
    /// <summary>
    /// WPF action that displays a folder picker dialog and returns the chosen folder
    /// in the string Value field of a <see cref="StringMessageContent"/>.
    /// </summary>
    class ChooseFolderAction : MessageActionBase
    {
        #region Overrides

        protected override void Invoke(object parameter)
        {
            MessageArgs<StringMessageContent> args = parameter as MessageArgs<StringMessageContent>;
            if (args != null)
            {
                args.Content.Confirmed = false;
                FolderBrowserDialog dlg = new FolderBrowserDialog();
                dlg.SelectedPath = args.Content.Value;
                dlg.ShowNewFolderButton = true;
                dlg.ShowDialog();
                if (!string.IsNullOrEmpty(dlg.SelectedPath))
                {
                    args.Content.Value = dlg.SelectedPath;
                    args.Content.Confirmed = true;
                    args.Respond();
                }
            }
            else
            {
                throw new InvalidOperationException("Expected to receive Message<StringMessageContent> as parameter.");
            }
        }

        /// <summary>
        /// Dummy implementation of the abstract method in the parent class.
        /// Will not be called because this class also overrides <see cref="Invoke"/>.
        /// </summary>
        /// <exception cref="InvalidOperationException">If this method is called (which
        /// it shouldn't, by design).</exception>
        /// <returns>Nothing.</returns>
        protected override Window CreateView()
        {
            throw new InvalidOperationException("This method should never be invoked in this class.");
        }

        #endregion
    }
}

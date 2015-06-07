/* FileFolderActionBase.cs
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
using Bovender.Mvvm.Messaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Abstract base class for the <see cref="ChooseFileSaveAction"/> and
    /// <see cref="ChooseFolderAction"/> classes.
    /// </summary>
    public abstract class FileFolderActionBase : MessageActionBase
    {
        #region Public properties

        public string Description { get; set; }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Displays an dialog to choose file or folder names.
        /// </summary>
        /// <param name="defaultString">Indicates the default
        /// path and/or file name and/or extension.</param>
        /// <returns>Valid file name/path, or empty string if
        /// the dialog was cancelled.</returns>
        protected abstract string GetDialogResult(
            string defaultString,
            string filter);

        #endregion

        #region Protected properties

        protected StringMessageContent MessageContent { get; private set; }

        #endregion

        #region Overrides

        protected override void Invoke(object parameter)
        {
            MessageArgs<FileNameMessageContent> args = parameter as MessageArgs<FileNameMessageContent>;
            MessageContent = args.Content;
            string result = GetDialogResult(args.Content.Value, args.Content.Filter);
            args.Content.Confirmed = !string.IsNullOrEmpty(result);
            if (args.Content.Confirmed)
            {
                args.Content.Value = result;
            };
            args.Respond();
        }

        protected override System.Windows.Window CreateView()
        {
            throw new InvalidOperationException(
                "This class does not create WPF views and this method should never be called.");
        }

        #endregion
    }
}

/* FileDialogActionBase.cs
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
using System.Windows.Forms;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// Base class for actions that use dialogs based on
    /// System.Windows.Forms.FileDialog.
    /// </summary>
    public abstract class FileDialogActionBase : FileFolderActionBase
    {
        #region Public properties

        /// <summary>
        /// Filter string for the file dialog. Filter entries consist
        /// of a description and an extensions, separated by a pipe
        /// symbol; multiple entries are separated by pipe symbols as
        /// well:
        /// "Image Files(*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|All files (*.*)|*.*"
        /// </summary>
        /// <remarks>
        /// This property will only be used to set the dialog filter
        /// if the derived class does not set the property in its implementation
        /// of the abstract GetDialog() method.
        /// </remarks>
        public string Filter { get; set; }

        #endregion

        #region Abstract methods

        protected abstract FileDialog GetDialog(
            string defaultString,
            string filter);

        #endregion

        #region Overrides

        protected override string GetDialogResult(
            string defaultString,
            string filter)
        {
            FileDialog dlg = GetDialog(defaultString, filter);
            dlg.Title = this.Caption;
            if (String.IsNullOrEmpty(dlg.Filter))
            {
                if (!String.IsNullOrEmpty(Filter))
                {
                    dlg.Filter = Filter;
                }
                else
                {
                    dlg.Filter = _messageContent.Filter;
                }
            }
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                return dlg.FileName;
            }
            else
            {
                return String.Empty;
            }
        }

        protected override void Invoke(object parameter)
        {
            MessageArgs<FileNameMessageContent> args = parameter as MessageArgs<FileNameMessageContent>;
            if (args == null)
            {
                throw new ArgumentException(
                    "Expected Message with FileNameMessageContent, not " +
                    parameter.GetType().ToString());
            }
            _messageContent = args.Content;
            string result = GetDialogResult(args.Content.Value, args.Content.Filter);
            args.Content.Confirmed = !string.IsNullOrEmpty(result);
            if (args.Content.Confirmed)
            {
                args.Content.Value = result;
            };
            args.Respond();
        }

        #endregion

        #region Private fields

        private FileNameMessageContent _messageContent;

        #endregion
    }
}

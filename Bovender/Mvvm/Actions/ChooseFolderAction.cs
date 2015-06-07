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
using System.Windows.Forms;
using Bovender.Mvvm.Messaging;

namespace Bovender.Mvvm.Actions
{
    /// <summary>
    /// MVVM action that queries the user for a folder.
    /// </summary>
    /// <remarks>
    /// To be used with MVVM messages that carry a <see cref="StringMessageContent"/>.
    /// </remarks>
    public class ChooseFolderAction : FileFolderActionBase
    {
        protected override string GetDialogResult(
            string defaultString,
            string filter)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.SelectedPath = defaultString;
            dlg.ShowNewFolderButton = true;
            dlg.Description = Description;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                return dlg.SelectedPath;
            }
            else
            {
                return String.Empty;
            }
        }
    }
}

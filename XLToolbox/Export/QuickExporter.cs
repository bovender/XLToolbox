/* QuickExporter.cs
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
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.Actions;
using XLToolbox.Export.Models;
using XLToolbox.Export.ViewModels;
using XLToolbox.Excel.Instance;

namespace XLToolbox.Export
{
    /// <summary>
    /// Provides user-entry points for the 'quick' export
    /// functions that re-use previously used settings.
    /// </summary>
    public class QuickExporter
    {
        #region Public methods

        /// <summary>
        /// Exports the current selection using the last settings, if available.
        /// </summary>
        public void ExportSelection()
        {
            Preset p = Preset.FromLastUsed(
                ExcelInstance.Application.ActiveWorkbook);
            if (p == null)
            {
                Dispatcher.Execute(Command.ExportSelection);
            }
            else
            {
                SingleExportSettingsViewModel svm = new SingleExportSettingsViewModel(p);
                svm.ChooseFileNameMessage.Sent += ChooseFileNameMessage_Sent; 
                if (svm.ChooseFileNameCommand.CanExecute(null))
                {
                    svm.ChooseFileNameCommand.Execute(null);
                }
            }
        }

        /// <summary>
        /// Performs a batch export using the last used settings, if available.
        /// </summary>
        public void ExportBatch()
        {
            BatchExportSettingsViewModel bvm = BatchExportSettingsViewModel.FromLastUsed(
                ExcelInstance.Application.ActiveWorkbook);
            if ((bvm != null) && bvm.ChooseFolderCommand.CanExecute(null))
            {
                bvm.ChooseFolderMessage.Sent += ChooseFolderMessage_Sent;
                bvm.ExportProcessMessage.Sent +=
                    (sender, args) =>
                    {
                        ProcessAction a = new ProcessAction();
                        a.Caption = Strings.BatchExport;
                        a.CancelButtonText = Strings.Cancel;
                        a.Invoke(args);
                    };
                bvm.ChooseFolderCommand.Execute(null);
            }
            else
            {
                if (bvm != null)
                {
                    bvm = new BatchExportSettingsViewModel();
                    // Do not 'sanitize' the export options, so that the user
                    // can see the selected, but disabled options.
                    bvm.InjectInto<Views.BatchExportSettingsView>().ShowDialog();
                }
                else
                {
                    Dispatcher.Execute(Command.BatchExport);
                }
            }
        }

        #endregion

        #region Private methods

        void ChooseFileNameMessage_Sent(object sender, MessageArgs<FileNameMessageContent> e)
        {
            ChooseFileSaveAction action = new ChooseFileSaveAction();
            action.Invoke(e);
        }

        void ChooseFolderMessage_Sent(object sender, MessageArgs<StringMessageContent> e)
        {
            ChooseFolderAction action = new ChooseFolderAction();
            action.Invoke(e);
        }

        #endregion
    }
}

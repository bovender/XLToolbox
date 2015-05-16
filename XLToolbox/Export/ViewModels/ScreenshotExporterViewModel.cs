/* ScreenshotExporterViewModel.cs
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
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;
using XLToolbox.Export;
using XLToolbox.Export.Models;
using Xl = Microsoft.Office.Interop.Excel;

namespace XLToolbox.Export.ViewModels
{
    /// <summary>
    /// View model for the <see cref="ScreenshotExporter"/> class.
    /// Provides commands in accordance with the MVVM pattern used
    /// by Bovender.
    /// </summary>
    public class ScreenshotExporterViewModel : ViewModelBase
    {
        #region Commands

        public DelegatingCommand ExportSelectionCommand
        {
            get
            {
                if (_exportSelectionCommand == null)
                {
                    _exportSelectionCommand = new DelegatingCommand(
                        (param) => DoChooseFileName(),
                        (param) => CanExportSelection()
                        );
                }
                return _exportSelectionCommand;
            }
        }

        #endregion

        #region Messages

        public Message<FileNameMessageContent> ChooseFileNameMessage
        {
            get
            {
                if (_chooseFileNameMessage == null)
                {
                    _chooseFileNameMessage = new Message<FileNameMessageContent>();
                }
                return _chooseFileNameMessage;
            }
        }
        #endregion

        #region Private methods

        private void DoChooseFileName()
        {
            if (CanExportSelection())
            {
                string defaultPath = Properties.Settings.Default.ExportPath;
                WorkbookStorage.Store store = new WorkbookStorage.Store();
                string path = store.Get(
                    Properties.StoreNames.Default.ExportPath, defaultPath);
                ChooseFileNameMessage.Send(
                    new FileNameMessageContent(
                        path,
                        FileType.Png.ToFileFilter()),
                    DoExportSelection);
            }
        }

        private void DoExportSelection(FileNameMessageContent messageContent)
        {
            if (CanExportSelection() && messageContent.Confirmed)
            {
                WorkbookStorage.Store store = new WorkbookStorage.Store();
                store.Put(Properties.StoreNames.Default.ExportPath, messageContent.Value);
                Properties.Settings.Default.ExportPath =
                    Bovender.PathHelpers.GetDirectoryPart(messageContent.Value);
                ScreenshotExporter exporter = new ScreenshotExporter();
                exporter.ExportSelection(messageContent.Value);
            }
        }

        private bool CanExportSelection()
        {
            Xl.Application app = Excel.ViewModels.Instance.Default.Application;
            return !(app == null || app.Selection is Xl.Range);
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            return null;
        }

        #endregion

        #region Private properties

        private WorkbookStorage.Store Store
        {
            get
            {
                if (_store == null)
                {
                    _store = new WorkbookStorage.Store();
                }
                return _store;
            }
        }

        #endregion

        #region Private fields

        private DelegatingCommand _exportSelectionCommand;
        private Message<FileNameMessageContent> _chooseFileNameMessage;
        private WorkbookStorage.Store _store;

        #endregion
    }
}

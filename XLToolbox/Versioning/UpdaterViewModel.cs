/* UpdaterViewModel.cs
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
using Bovender.Mvvm.Actions;
using Bovender.Extensions;

namespace XLToolbox.Versioning
{
    public class UpdaterViewModel : Bovender.Versioning.UpdaterViewModel
    {
        #region Constructor

        public UpdaterViewModel(Updater updater)
            : base(updater)
        {
            XLToolbox.Excel.ViewModels.Instance.Default.ShuttingDown += Instance_ShuttingDown;
            ShowProgressMessage.Sent += (sender2, args2) =>
            {
                args2.Content.InjectInto<XLToolbox.Versioning.UpdaterProcessView>().ShowInForm();
            };
            DownloadFinishedMessage.Sent += (sender2, args2) =>
            {
                Bovender.Mvvm.Actions.ProcessCompletedAction a = new Bovender.Mvvm.Actions.ProcessCompletedAction();
                a.Caption = XLToolbox.Strings.UpdateAvailable;
                a.Message = XLToolbox.Strings.UpdateHasBeenDownloaded;
                a.OkButtonText = XLToolbox.Strings.OK;
                a.InvokeWithContent(args2.Content);
            };
            DownloadFailedMessage.Sent += (sender2, args2) =>
            {
                Bovender.Mvvm.Actions.ProcessCompletedAction a = new Bovender.Mvvm.Actions.ProcessCompletedAction();
                a.Caption = XLToolbox.Strings.UpdateAvailable;
                a.Message = XLToolbox.Strings.ErrorOccurredWhileDownloading;
                a.OkButtonText = XLToolbox.Strings.OK;
                a.InvokeWithContent(args2.Content);
            };
        }

        #endregion

        #region Public method

        /// <summary>
        /// Injects this view model into an UpdateAvailableView and shows it as a dialog.
        /// </summary>
        public void ShowUpdateAvailableView()
        {
            DownloadFolder = UserSettings.UserSettings.Default.DownloadFolder;
            InjectInto<UpdateAvailableView>().ShowInForm();
            UserSettings.UserSettings.Default.DownloadFolder = DownloadFolder;
        }

        #endregion

        #region Event handlers

        private void Instance_ShuttingDown(object sender, Excel.ViewModels.InstanceShutdownEventArgs e)
        {
            CancelProcess();
            XLToolbox.Excel.ViewModels.Instance.Default.ShuttingDown -= Instance_ShuttingDown;
        }
 
        #endregion
    }
}

/* ThisAddIn.cs
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
using Bovender.Mvvm;
using Bovender.Unmanaged;
using Bovender.Versioning;
using System;
using XLToolbox.Excel.Instance;
using XLToolbox.ExceptionHandler;
using XLToolbox.Mvvm.Views;
using XLToolbox.Greeter;
using Threading = System.Windows.Threading;

namespace XLToolbox
{
    public partial class ThisAddIn : IDisposable
    {
        #region Private fields

        private Threading.Dispatcher _dispatcher;
        private DllManager _dllManager;
        private UpdaterViewModel _updaterVM;

        #endregion

        #region Startup/Shutdown

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Get a hold of the current dispatcher so we can create an
            // update notification window from a different thread
            // when checking for updates.
            _dispatcher = Threading.Dispatcher.CurrentDispatcher;

            // Make the current Excel instance globally available
            // even for the non-VSTO components of this addin
            ExcelInstance.Application = Globals.ThisAddIn.Application;

            Bovender.ExceptionHandler.CentralHandler.ManageExceptionCallback += CentralHandler_ManageExceptionCallback;

            // Distract the user :-)
            MaybeCheckForUpdate();
            GreetUser();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Use && to perform lazy evaluation
            if (_updaterVM != null && _updaterVM.InstallUpdateCommand.CanExecute(null))
            {
                _updaterVM.InjectInto<XLToolbox.Versioning.InstallUpdateView>().ShowDialog();
            };
        }

        #endregion

        #region Private methods

        private void GreetUser()
        {
            SemanticVersion lastSeenVersion = new SemanticVersion(
                Properties.Settings.Default.LastVersionSeen);
            SemanticVersion currentVersion = XLToolbox.Versioning.SemanticVersion.CurrentVersion();
            if (currentVersion > lastSeenVersion)
            {
                Workarounds.ShowModelessInExcel<GreeterView>(new GreeterViewModel());
                Properties.Settings.Default.LastVersionSeen = currentVersion.ToString();
                Properties.Settings.Default.Save();
            }
        }

        /// <summary>
        /// Shows an error message to the user when the central exception handler
        /// (implemented in Bovender.dll) raises the ManageExceptionCallback event.
        /// </summary>
        /// <param name="sender">Object where the exception occurred.</param>
        /// <param name="e">Instance of ManageExceptionEventArgs with additional information.</param>
        void CentralHandler_ManageExceptionCallback(object sender, Bovender.ExceptionHandler.ManageExceptionEventArgs e)
        {
            e.IsHandled = true;
            ExceptionViewModel vm = new ExceptionViewModel(e.Exception);
            vm.InjectInto<ExceptionView>().ShowDialog();
        }

        /// <summary>
        /// Performs an online update check, but only if the specified number of
        /// days between update checks has passed.
        /// </summary>
        private void MaybeCheckForUpdate()
        {
            DateTime lastCheck = Properties.Settings.Default.LastUpdateCheck;
            DateTime today = DateTime.Today;
            if ((today - lastCheck).Days >= Properties.Settings.Default.UpdateCheckInterval)
            {
                _updaterVM = new UpdaterViewModel(new XLToolbox.Versioning.Updater());
                if (_updaterVM.IsUserAuthorized)
                {
                    _updaterVM.UpdateAvailableMessage.Sent += (sender, args) =>
                    {
                        Workarounds.ShowModelessInExcel<XLToolbox.Versioning.UpdateAvailableView>(_updaterVM);
                    };
                    _updaterVM.CheckForUpdateCommand.Execute(null);
                }
                Properties.Settings.Default.LastUpdateCheck = DateTime.Today;
                Properties.Settings.Default.Save();
            }
        }

        #endregion

        #region Ribbon

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

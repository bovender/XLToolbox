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
using System;
using Threading = System.Windows.Threading;
using System.Configuration;
using Bovender.Versioning;
using Bovender.Mvvm.Actions;
using Bovender.Extensions;
using Ver = XLToolbox.Versioning;
using XLToolbox.Excel.ViewModels;
using XLToolbox.ExceptionHandler;
using XLToolbox.Greeter;
using System.Diagnostics;
using System.Windows.Threading;

namespace XLToolboxForExcel
{
    public partial class ThisAddIn : IDisposable
    {
        #region Startup/Shutdown

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Delete user config file that may be left over from NG developmental
            // versions. We don't need it anymore and it causes nasty crashes.
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
                if (System.IO.File.Exists(config.FilePath))
                {
                    System.IO.File.Delete(config.FilePath);
                }
            }
            catch { }

            // Get a hold of the current dispatcher so we can create an
            // update notification window from a different thread
            // when checking for updates.
            _dispatcher = Threading.Dispatcher.CurrentDispatcher;

            // Make the current Excel instance globally available
            // even for the non-VSTO components of this addin
            Instance.Default = new Instance(Globals.ThisAddIn.Application);
            Ribbon.ExcelApp = Instance.Default.Application;

            // Register Excel's main window handle to facilitate interop with WPF.
            Bovender.Extensions.WindowExtensions.TopLevelWindow = (IntPtr)Globals.ThisAddIn.Application.Hwnd;

            // Make the CustomTaskPanes available for dispatcher methods.
            XLToolbox.Globals.CustomTaskPanes = CustomTaskPanes;

            Bovender.ExceptionHandler.CentralHandler.ManageExceptionCallback += CentralHandler_ManageExceptionCallback;
            Bovender.WpfHelpers.RegisterTextBoxSelectAll();
            Bovender.ExceptionHandler.CentralHandler.DumpFile =
                System.IO.Path.Combine(System.IO.Path.GetTempPath() + Properties.Settings.Default.DumpFile);
            AppDomain.CurrentDomain.UnhandledException += Bovender.ExceptionHandler.CentralHandler.AppDomain_UnhandledException;

            PerformSanityChecks();
            MaybeCheckForUpdate();
            GreetUser();

            XLToolbox.Keyboard.Manager.Default.EnableShortcuts();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            XLToolbox.UserSettings.Default.Save();
            Bovender.Versioning.UpdaterViewModel uvm = Ver.UpdaterViewModel.Instance;
            if (uvm.IsUpdatePending && uvm.InstallUpdateCommand.CanExecute(null))
            {
                // Must show the InstallUpdateView modally, because otherwise Excel would
                // continue to shut down and immediately remove the view while doing so.
                uvm.InjectInto<XLToolbox.Versioning.InstallUpdateView>().ShowDialogInForm();
            };

            // Prevent "LocalDataSlot storage has been freed" exceptions;
            // see http://j.mp/localdatastoreslot
            Dispatcher.CurrentDispatcher.InvokeShutdown();
        }

        #endregion

        #region Properties

        public Ribbon Ribbon
        {
            get
            {
                if (_ribbon == null)
                {
                    _ribbon = new Ribbon();
                }
                return _ribbon;
            }
        }

        #endregion

        #region Private methods

        private void GreetUser()
        {
            SemanticVersion lastVersionSeen = new SemanticVersion(
                XLToolbox.UserSettings.Default.LastVersionSeen);
            SemanticVersion currentVersion = XLToolbox.Versioning.SemanticVersion.CurrentVersion();
            if (currentVersion > lastVersionSeen)
            {
                GreeterViewModel gvm = new GreeterViewModel();
                gvm.InjectAndShowInThread<GreeterView>();
                XLToolbox.UserSettings.Default.LastVersionSeen = currentVersion.ToString();
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
            vm.InjectInto<ExceptionView>().ShowDialogInForm();
        }

        /// <summary>
        /// Performs an online update check, but only if the specified number of
        /// days between update checks has passed.
        /// </summary>
        private void MaybeCheckForUpdate()
        {
            DateTime lastCheck = XLToolbox.UserSettings.Default.LastUpdateCheck;
            DateTime today = DateTime.Today;
            if ((today - lastCheck).Days >= XLToolbox.UserSettings.Default.UpdateCheckInterval)
            {
                _installUpdateView = new Ver.InstallUpdateView();
                UpdaterViewModel updaterVM = Ver.UpdaterViewModel.Instance;
                if (updaterVM.CanCheckForUpdate)
                {
                    updaterVM.UpdateAvailableMessage.Sent += (sender, args) =>
                    {
                        // Must show the view in a separate thread in order for it to
                        // receive keyboard input (otherwise, Excel would grab all keyboard
                        // events).
                        updaterVM.InjectAndShowInThread<Ver.UpdateAvailableView>();
                    };
                    updaterVM.CheckForUpdateCommand.Execute(null);
                }
                XLToolbox.UserSettings.Default.LastUpdateCheck = DateTime.Today;
            }
        }

        private void PerformSanityChecks()
        {
            XLToolbox.Legacy.LegacyToolbox.DeactivateObsoleteVbaAddin();
            XLToolbox.UserSettings userSettings = XLToolbox.UserSettings.Default;
            if (userSettings.Exception != null)
            {
                Bovender.UserSettings.UserSettingsExceptionViewModel vm =
                    new Bovender.UserSettings.UserSettingsExceptionViewModel(userSettings);
                vm.InjectInto<XLToolbox.Mvvm.Views.UserSettingsExceptionView>().ShowDialogInForm();
            }
        }

        #endregion

        #region Private fields

        private Threading.Dispatcher _dispatcher;
        private Ribbon _ribbon;
        private XLToolbox.Versioning.InstallUpdateView _installUpdateView;

        #endregion

        #region VBA API

        protected override object RequestComAddInAutomationService()
        {
            return XLToolbox.Vba.Api.Default;
        }

        #endregion

        #region Ribbon

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Ribbon;
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

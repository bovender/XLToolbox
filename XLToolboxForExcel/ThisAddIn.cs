/* ThisAddIn.cs
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
using Threading = System.Windows.Threading;
using System.Configuration;
using Bovender.Versioning;
using Bovender.Extensions;
using Ver = XLToolbox.Versioning;
using XLToolbox.Excel.ViewModels;
using XLToolbox.ExceptionHandler;
using XLToolbox.Greeter;
using System.Windows.Threading;
using NLog;

namespace XLToolboxForExcel
{
    public partial class ThisAddIn : IDisposable
    {
        #region Startup/Shutdown

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Delete user config file that may be left over from NG developmental
            // versions. We don't need it anymore and it causes nasty crashes.
            // Must do this before using NLog!
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
                if (System.IO.File.Exists(config.FilePath))
                {
                    System.IO.File.Delete(config.FilePath);
                }
            }
            catch { }

#if DEBUG
            XLToolbox.Logging.LogFile.Default.EnableDebugLogging();
#endif

            Logger.Info("Begin startup");

            // Get a hold of the current dispatcher so we can create an
            // update notification window from a different thread
            // when checking for updates.
            _dispatcher = Threading.Dispatcher.CurrentDispatcher;

            // Make the current Excel instance globally available
            // even for the non-VSTO components of this addin
            Instance.Default = new Instance(Globals.ThisAddIn.Application);
            Ribbon.ExcelApp = Instance.Default.Application;

            // Register Excel's main window handle to facilitate interop with WPF.
            _mainWindow = (IntPtr)Globals.ThisAddIn.Application.Hwnd;
            Bovender.Extensions.WindowExtensions.TopLevelWindow = _mainWindow;

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

            // Enable the keyboard shortcuts if no settings were previously saved,
            // i.e. if this appears to be the first run.
            if (!XLToolbox.UserSettings.UserSettings.Default.WasFromFile)
            {
                XLToolbox.Keyboard.Manager.Default.IsEnabled = true;
            }

            if (XLToolbox.UserSettings.UserSettings.Default.SheetManagerVisible)
            {
                XLToolbox.SheetManager.SheetManagerPane.Default.Visible = true;
            }

            Logger.Info("Finished startup");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Logger.Info("Begin shutdown");
            XLToolbox.UserSettings.UserSettings.Default.Running = false;
            XLToolbox.UserSettings.UserSettings.Default.Save();

            if (XLToolbox.Legacy.LegacyToolbox.IsInitialized)
            {
                Logger.Info("Disposing legacy add-in");
                XLToolbox.Legacy.LegacyToolbox.Default.Dispose();
            }

            Bovender.Versioning.UpdaterViewModel uvm = Ver.UpdaterViewModel.Instance;
            if (uvm.IsUpdatePending && uvm.InstallUpdateCommand.CanExecute(null))
            {
                // Must show the InstallUpdateView modally, because otherwise Excel would
                // continue to shut down and immediately remove the view while doing so.
                uvm.InjectInto<XLToolbox.Versioning.InstallUpdateView>().ShowDialogInForm();
            };

            Ribbon.PrepareShutdown();

            // Prevent "LocalDataSlot storage has been freed" exceptions;
            // see http://j.mp/localdatastoreslot
            Dispatcher.CurrentDispatcher.InvokeShutdown();
            Logger.Info("Finish shutdown");
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
                XLToolbox.UserSettings.UserSettings.Default.LastVersionSeen);
            SemanticVersion currentVersion = XLToolbox.Versioning.SemanticVersion.CurrentVersion();
            Logger.Info("Current version: {0}; last was {1}", currentVersion, lastVersionSeen);
            if (currentVersion > lastVersionSeen)
            {
                Logger.Info("Greeting user");
                GreeterViewModel gvm = new GreeterViewModel();
                gvm.InjectAndShowDialogInThread<GreeterView>(_mainWindow);
                XLToolbox.UserSettings.UserSettings.Default.LastVersionSeen = currentVersion.ToString();
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
            Logger.Fatal("Central exception hander callback: {0}", e);
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
            DateTime lastCheck = XLToolbox.UserSettings.UserSettings.Default.LastUpdateCheck;
            DateTime today = DateTime.Today;
            if ((today - lastCheck).Days >= XLToolbox.UserSettings.UserSettings.Default.UpdateCheckInterval)
            {
                _installUpdateView = new Ver.InstallUpdateView();
                UpdaterViewModel updaterVM = Ver.UpdaterViewModel.Instance;
                if (updaterVM.CanCheckForUpdate)
                {
                    Logger.Info("Checking for update");
                    updaterVM.UpdateAvailableMessage.Sent += (sender, args) =>
                    {
                        // Must show the view in a separate thread in order for it to
                        // receive keyboard input (otherwise, Excel would grab all keyboard
                        // events).
                        updaterVM.InjectAndShowInThread<Ver.UpdateAvailableView>();
                    };
                    _dispatcher.BeginInvoke(new Action(() => updaterVM.CheckForUpdateCommand.Execute(null)));
                }
                XLToolbox.UserSettings.UserSettings.Default.LastUpdateCheck = DateTime.Today;
            }
        }

        private void PerformSanityChecks()
        {
            XLToolbox.UserSettings.UserSettings userSettings = XLToolbox.UserSettings.UserSettings.Default;
            Logger.Info("Performing sanity checks");

            // Deactivating the VBA add-in can cause crashes; we now do it in the installer
            // XLToolbox.Legacy.LegacyToolbox.DeactivateObsoleteVbaAddin();

            if (userSettings.Running)
            {
                XLToolbox.Logging.LogFileViewModel vm = new XLToolbox.Logging.LogFileViewModel();
                if (userSettings.EnableLogging)
                {
                    vm.InjectInto<XLToolbox.Logging.IncompleteShutdownLoggingEnabled>().ShowDialogInForm();
                }
                else
                {
                    vm.InjectInto<XLToolbox.Logging.IncompleteShutdownLoggingDisabled>().ShowDialogInForm();
                }
            }
            if (userSettings.Exception != null)
            {
                Bovender.UserSettings.UserSettingsExceptionViewModel vm =
                    new Bovender.UserSettings.UserSettingsExceptionViewModel(userSettings);
                vm.InjectInto<XLToolbox.Mvvm.Views.UserSettingsExceptionView>().ShowDialogInForm();
            }
            userSettings.Running = true;
            userSettings.Save();
            Logger.Info("Sanity checks completed");
        }

        #endregion

        #region Private fields

        private Threading.Dispatcher _dispatcher;
        private Ribbon _ribbon;
        private XLToolbox.Versioning.InstallUpdateView _installUpdateView;
        private IntPtr _mainWindow;

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

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

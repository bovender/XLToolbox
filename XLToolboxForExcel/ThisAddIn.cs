/* ThisAddIn.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using System.Configuration;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using NLog;
using Bovender.Extensions;
using Bovender.Mvvm.Actions;
using XLToolbox.Excel.ViewModels;
using XLToolbox.ExceptionHandler;
using XLToolbox.Greeter;
using XLToolbox.Versioning;

namespace XLToolboxForExcel
{
    [ProgId("DanielsXLToolbox")]
    [Guid("8DDC0086-3BAB-4D31-B5FD-6DEE3A1C78C9")]
    public partial class ThisAddIn : IDisposable
    {
        #region Startup/Shutdown

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
#if DEBUG
            XLToolbox.Logging.LogFile.Default.EnableDebugLogging();
#endif

            Logger.Info("ThisAddIn_Startup: Begin startup");

            // Register Excel's main window handle to facilitate interop with WPF.
            Bovender.Win32Window.MainWindowHandleProvider =
                new Func<IntPtr>(() => (IntPtr)Globals.ThisAddIn.Application.Hwnd);

            Bovender.Unmanaged.DllManager.AlternativeDir = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    XLToolbox.Properties.Settings.Default.AppDataFolder
                    );

            // Get a hold of the current dispatcher so we can create an
            // update notification window from a different thread
            // when checking for updates.
            _dispatcher = Dispatcher.CurrentDispatcher;
            Bovender.WpfHelpers.MainDispatcher = _dispatcher;
            Updater.CanCheck = true;

            // Make the current Excel instance globally available
            // even for the non-VSTO components of this addin
            Instance.Default = new Instance(Globals.ThisAddIn.Application);
            Ribbon.ExcelApp = Instance.Default.Application;

            // Make the CustomTaskPanes available for dispatcher methods.
            XLToolbox.Globals.CustomTaskPanes = CustomTaskPanes;

            Bovender.ExceptionHandler.CentralHandler.ManageExceptionCallback += CentralHandler_ManageExceptionCallback;
            Bovender.WpfHelpers.RegisterTextBoxSelectAll();
            Bovender.ExceptionHandler.CentralHandler.DumpFile =
                System.IO.Path.Combine(System.IO.Path.GetTempPath() + Properties.Settings.Default.DumpFile);
            AppDomain.CurrentDomain.UnhandledException += Bovender.ExceptionHandler.CentralHandler.AppDomain_UnhandledException;

            Ribbon.SubscribeToEvents();
            PerformSanityChecks();

            XLToolbox.Backup.Backups.BackupFailed += Backups_BackupFailed;

            if (XLToolbox.UserSettings.UserSettings.Default.SheetManagerVisible)
            {
                XLToolbox.SheetManager.TaskPaneManager.Default.Visible = true;
            }

            TestDllAvailability();
            if (!GreetUser()) MaybeCheckForUpdate();
            Logger.Info("ThisAddIn_Startup: Done");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Logger.Info("ThisAddIn_Shutdown: Starting to clean up after ourselves");
            XLToolbox.UserSettings.UserSettings.Default.Running = false;
            XLToolbox.UserSettings.UserSettings.Default.Save();
            XLToolbox.Backup.Backups.BackupFailed -= Backups_BackupFailed;

            if (XLToolbox.Legacy.LegacyToolbox.IsInitialized)
            {
                Logger.Info("Disposing legacy add-in");
                XLToolbox.Legacy.LegacyToolbox.Default.Dispose();
            }

            if (Updater.Default != null && Updater.Default.Status == Bovender.Versioning.UpdaterStatus.Downloaded)
            {
                Logger.Info("ThisAddIn_Shutdown: Offering to install update that was previously downloaded");
                UpdaterViewModel updaterVM = new UpdaterViewModel(Updater.Default);
                // Must show the InstallUpdateView modally, because otherwise Excel would
                // continue to shut down and immediately remove the view while doing so.
                updaterVM.InjectInto<XLToolbox.Versioning.InstallUpdateView>().ShowDialogInForm();
            };

            // Prevent "LocalDataSlot storage has been freed" exceptions;
            // see http://j.mp/localdatastoreslot
            Dispatcher.CurrentDispatcher.InvokeShutdown();
            Logger.Info("ThisAddIn_Shutdown: Done.");
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

        private bool GreetUser()
        {
            bool result = false;
            SemanticVersion lastVersionSeen = new SemanticVersion(
                XLToolbox.UserSettings.UserSettings.Default.LastVersionSeen);
            Logger.Info("GreetUser: Current version: {0}; last was {1}", SemanticVersion.Current, lastVersionSeen);
            if (SemanticVersion.Current > lastVersionSeen)
            {
                System.Threading.Timer timer = new System.Threading.Timer(
                    (obj) => { 
                        _dispatcher.BeginInvoke((Action)(() =>
                        {
                            Logger.Info("GreetUser: showing welcome dialog");
                            XLToolbox.UserSettings.UserSettings.Default.LastVersionSeen = SemanticVersion.Current.ToString();
                            GreeterViewModel gvm = new GreeterViewModel();
                            gvm.InjectInto<GreeterView>().ShowInForm();
                        }));
                    },
                    null, 250, System.Threading.Timeout.Infinite);
                result = true;
            }
            return result;
        }

        /// <summary>
        /// Performs an online update check, but only if the specified number of
        /// days between update checks has passed and if the user has consented.
        /// </summary>
        private void MaybeCheckForUpdate()
        {
            if (XLToolbox.UserSettings.UserSettings.Default.EnableUpdateChecks == false)
            {
                Logger.Info("MaybeCheckForUpdate: Update checks not enabled by user.");
                return;
            }
            DateTime lastCheck = XLToolbox.UserSettings.UserSettings.Default.LastUpdateCheck;
            DateTime today = DateTime.Today;
            if ((today - lastCheck).Days >= XLToolbox.UserSettings.UserSettings.Default.UpdateCheckInterval)
            {
                ReleaseInfoViewModel releaseInfoVM = new ReleaseInfoViewModel();
                releaseInfoVM.ProcessFinishedMessage.Sent += (sender, args) =>
                {
                    if (releaseInfoVM.Status == Bovender.Versioning.ReleaseInfoStatus.InfoAvailable)
                    {
                        Logger.Info("MaybeCheckForUpdate: Info available, remembering current date");
                        XLToolbox.UserSettings.UserSettings.Default.LastUpdateCheck = DateTime.Today;
                    }
                };
                Logger.Info("MaybeCheckForUpdate: Checking for update");
                Updater.CanCheck = false;
                releaseInfoVM.StartProcess();
            }
        }

        private void PerformSanityChecks()
        {
            XLToolbox.UserSettings.UserSettings userSettings = XLToolbox.UserSettings.UserSettings.Default;
            Logger.Info("Performing sanity checks");
            ExceptionViewModel evm = new ExceptionViewModel(null);
            Logger.Info("+++ Excel version:    {0}, {1}", evm.ExcelVersion, evm.ExcelBitness);
            Logger.Info("+++ OS version:       {0}, {1}", evm.OS, evm.OSBitness);
            Logger.Info("+++ CLR version:      {0}, {1}", evm.CLR, evm.ProcessBitness);
            Logger.Info("+++ VSTOR version:    {0}", evm.VstoRuntime);
            Logger.Info("+++ Bovender version: {0}", evm.BovenderFramework);

            // Deactivating the VBA add-in can cause crashes; we now do it in the installer
            // XLToolbox.Legacy.LegacyToolbox.DeactivateObsoleteVbaAddin();

            if (userSettings.Running)
            {
                XLToolbox.Logging.LogFileViewModel vm = new XLToolbox.Logging.LogFileViewModel();
                if (userSettings.EnableLogging)
                {
                    vm.InjectInto<XLToolbox.Logging.IncompleteShutdownLoggingEnabled>().ShowInForm();
                }
                else
                {
                    vm.InjectInto<XLToolbox.Logging.IncompleteShutdownLoggingDisabled>().ShowInForm();
                }
                userSettings.SheetManagerVisible = false;
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

        /// <summary>
        /// Shows an error message to the user when the central exception handler
        /// (implemented in Bovender.dll) raises the ManageExceptionCallback event.
        /// </summary>
        /// <param name="sender">Object where the exception occurred.</param>
        /// <param name="e">Instance of ManageExceptionEventArgs with additional information.</param>
        private void CentralHandler_ManageExceptionCallback(object sender, Bovender.ExceptionHandler.ManageExceptionEventArgs e)
        {
            Logger.Error(e.Exception);
            e.IsHandled = true;
            ExceptionViewModel vm = new ExceptionViewModel(e.Exception);
            vm.InjectInto<ExceptionView>().ShowDialogInForm();
        }

        private void Backups_BackupFailed(object sender, Bovender.ExceptionHandler.ManageExceptionEventArgs e)
        {
            XLToolbox.Backup.BackupFailedViewModel vm = new XLToolbox.Backup.BackupFailedViewModel(e.Exception);
            if (!vm.Suppress)
            {
                Logger.Info("Backups_BackupFailed: Informing user about failed backup");
                vm.InjectInto<XLToolbox.Backup.BackupFailedView>().ShowDialogInForm();
            }
            else
            {
                Logger.Info("Backups_BackupFailed: Failure message suppressed by user");
            }
        }

        /// <summary>
        /// Tests the availability of the FreeImage DLL and issues a warning
        /// to the user if the DLL is not available.
        /// </summary>
        private void TestDllAvailability()
        {
            using (Bovender.Unmanaged.DllManager dllManager = new Bovender.Unmanaged.DllManager())
            {
                try
                {
                    Logger.Info("TestDllAvailability: Testing freeimage.dll");
                    dllManager.LoadDll("freeimage.dll");
                }
                catch (Exception e)
                {
                    Logger.Warn("TestDllAvailability: Failed to load DLL");
                    Logger.Warn(e);
                    Ribbon.IsGraphicExportEnabled = false;
                    if (!XLToolbox.UserSettings.UserSettings.Default.SuppressDllWarning)
                    {
                        Logger.Warn("TestDllAvailability: Showing warning message");
                        SuppressibleNotificationAction a = new SuppressibleNotificationAction();
                        Bovender.Mvvm.Messaging.SuppressibleMessageContent mc = new Bovender.Mvvm.Messaging.SuppressibleMessageContent();
                        mc.OkButtonText = XLToolbox.Strings.OK;
                        mc.SuppressMessageText = XLToolbox.Strings.DoNotShowThisMessageAgain;
                        mc.Caption = XLToolbox.Strings.DllNotAvailableCaption;
                        mc.Message = XLToolbox.Strings.DllNotAvailableMessage;
                        a.InvokeWithContent(mc);
                        XLToolbox.UserSettings.UserSettings.Default.SuppressDllWarning = mc.Suppress;
                    }
                    else
                    {
                        Logger.Warn("TestDllAvailability: Warning message is suppressed by user");
                    }
                }
            }
        }

        #endregion

        #region Private fields

        private Dispatcher _dispatcher;
        private Ribbon _ribbon;

        #endregion

        #region VBA API

        protected override object RequestComAddInAutomationService()
        {
            Logger.Info("RequestComAddInAutomationService");
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

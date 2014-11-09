using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Threading;
using Threading = System.Windows.Threading;
using XLToolbox.Excel.Instance;
using Bovender.Versioning;
using Bovender.Unmanaged;
using Bovender.Mvvm;
using XLToolbox.Mvvm.ViewModels;
using XLToolbox.Mvvm.Views;
using XLToolbox.ExceptionHandler;

namespace XLToolbox
{
    public partial class ThisAddIn : IDisposable
    {
        #region Public properties

        public XLToolbox.Versioning.Updater Updater { get; set; }

        #endregion

        #region Private fields

        private Threading.Dispatcher _dispatcher;
        private DllManager _dllManager;

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

            // Load the FreeImage DLL
            _dllManager = new DllManager();
            _dllManager.LoadDll("FreeImage");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Use && to perform lazy evaluation
            if (Updater != null && Updater.Downloaded)
            {
                MessageBoxResult r = MessageBox.Show(Strings.UpdateWillBeInstalledNow,
                    Strings.UpdateAvailable, MessageBoxButton.OKCancel, MessageBoxImage.Information);
                if (r == MessageBoxResult.OK)
                {
                    Updater.InstallUpdate();
                }
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
                GreeterViewModel greeter = new GreeterViewModel();
                Workarounds.ShowModelessInExcel<GreeterView>(greeter);
                Properties.Settings.Default.LastVersionSeen = currentVersion.ToString();
                Properties.Settings.Default.Save();
            }
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
                Updater = new XLToolbox.Versioning.Updater();
                if (Updater.IsAuthorized)
                {
                    Updater.UpdateAvailable += Updater_UpdateAvailable;
                    Updater.NoUpdateAvailable += Updater_NoUpdateAvailable;
                    Updater.FetchVersionInformation();
                }
            }
        }

        void CentralHandler_ManageExceptionCallback(object sender, Bovender.ExceptionHandler.ManageExceptionEventArgs e)
        {
            e.IsHandled = true;
            ExceptionViewModel vm = new ExceptionViewModel(e.Exception);
            vm.InjectInto<ExceptionView>().ShowDialog();
        }

        #endregion

        #region Updates

        void Updater_NoUpdateAvailable(object sender, UpdateAvailableEventArgs e)
        {
            RememberUpdateCheck();
            Updater = null;
        }

        /// <summary>
        /// Is called when a new update is available.
        /// </summary>
        /// <param name="sender">Instance of Updater.</param>
        /// <param name="e">Relevant arguments.</param>
        void Updater_UpdateAvailable(object sender, UpdateAvailableEventArgs e)
        {
            RememberUpdateCheck();
            /* For thread-safe execution, we must defer the update notification
             * to the UI thread.
             */
            _dispatcher.Invoke(new System.Action(ShowUpdateInformation));
        }

        /// <summary>
        /// Thread-safe method that notifies the user of an available update.
        /// </summary>
        void ShowUpdateInformation()
        {
            Thread thread = new Thread(() =>
            {
                // TODO: Use MVVM here
                // WindowUpdateAvailable w = new WindowUpdateAvailable(Updater);
                // w.Show();
                // w.Closed += (sender2, e2) => w.Dispatcher.InvokeShutdown();
                // System.Windows.Threading.Dispatcher.Run();
            });
            thread.SetApartmentState( ApartmentState.STA);
            thread.Start();
        }

        /// <summary>
        /// Save the current date as the last date an update check was performed.
        /// </summary>
        /// <remarks>
        /// This method is only called by the event handlers Updater_NoUpdateAvailable
        /// and Updater_UpdateAvailable. If checking for updates failed, neither of these
        /// will be called, and we can check again whenever the add-in is started again.
        /// </remarks>
        void RememberUpdateCheck()
        {
                Properties.Settings.Default.LastUpdateCheck = DateTime.Today;
                Properties.Settings.Default.Save();
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using XLToolbox.Version;
using System.Threading;
using Threading = System.Windows.Threading;

namespace XLToolbox
{
    public partial class ThisAddIn
    {
        public Updater Updater { get; set; }
        private Threading.Dispatcher _dispatcher;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
#if !DEBUG
            Globals.Ribbons.Ribbon.GroupDebug.Visible = false;
#endif
            // Get a hold of the current dispatcher so we can create an
            // update notification window from a different thread
            // when checking for updates.
            _dispatcher = Threading.Dispatcher.CurrentDispatcher;

            MaybeCheckForUpdate();
            GreetUser();
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
            }
        }

        private void GreetUser()
        {
            SemanticVersion lastSeenVersion = new SemanticVersion(
                Properties.Settings.Default.LastVersionSeen);
            SemanticVersion currentVersion = SemanticVersion.CurrentVersion();
            if (currentVersion > lastSeenVersion)
            {
                WpfHelpers.ShowModelessInExcel<WindowGreeter>();
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
                Updater = new Updater();
                if (Updater.IsAuthorized)
                {
                    Updater.UpdateAvailable += Updater_UpdateAvailable;
                    Updater.NoUpdateAvailable += Updater_NoUpdateAvailable;
                    Updater.FetchVersionInformation();
                }
            }
        }

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
                WindowUpdateAvailable w = new WindowUpdateAvailable(Updater);
                w.Show();
                w.Closed += (sender2, e2) => w.Dispatcher.InvokeShutdown();
                System.Windows.Threading.Dispatcher.Run();
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using XLToolbox.Version;

namespace XLToolbox
{
    /// <summary>
    /// Interaction logic for WindowCheckForUpdate.xaml
    /// </summary>
    public partial class WindowCheckForUpdate : Window
    {
        public WindowCheckForUpdate()
        {
            InitializeComponent();
            Updater updater = new Updater();
            updater.UpdateAvailable += updater_OnUpdateAvailable;
            updater.FetchingVersionFailed += updater_FetchingVersionFailed;
            updater.NoUpdateAvailable += updater_NoUpdateAvailable;
            updater.FetchVersionInformation();
        }

        void updater_NoUpdateAvailable(object sender, UpdateAvailableEventArgs e)
        {
            stopProgressBar();
            MessageBox.Show(Strings.YouHaveTheLatestVersion, Strings.CheckForUpdates,
                MessageBoxButton.OK, MessageBoxImage.Information);
            dispatchClose();
        }

        void updater_FetchingVersionFailed(object sender, System.Net.DownloadStringCompletedEventArgs e)
        {
            stopProgressBar();
            MessageBox.Show(String.Format(Strings.FetchingVersionInformationFailed, e.Error.Message),
                Strings.CheckForUpdates, MessageBoxButton.OK);
            dispatchClose();
        }

        private void updater_OnUpdateAvailable(object sender, UpdateAvailableEventArgs e)
        {
            stopProgressBar();
            showUpdateAvailable(sender as Updater);
            dispatchClose();
        }

        private void stopProgressBar()
        {
            Action stopProgressBar = delegate()
            {
                ProgressBar.IsIndeterminate = false;
            };
            this.Dispatcher.Invoke(new Action(stopProgressBar));
        }

        private void dispatchClose()
        {
            this.Dispatcher.Invoke(new Action(this.Close));
        }

        /// <summary>
        /// Thread-safe method to show the update information window; can
        /// be called from event handlers that run in non-UI threads.
        /// </summary>
        /// <param name="updater"></param>
        private void showUpdateAvailable(Updater updater)
        {
            Action action;
            if (updater.IsAuthorized)
            {
                action = delegate()
                {
                    (new WindowUpdateAvailable(updater)).Show();
                };
            }
            else
            {
                action = delegate()
                {
                    (new WindowNotAuthorizedForUpdate(updater)).Show();
                };
            }
            this.Dispatcher.Invoke(new Action(action));
        }
    }
}

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
            updater.FetchVersionInformation();
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
            MessageBox.Show(String.Format("New version available!\n{0}\n{1}\n{2}",
                e.NewVersion.ToString(), e.DownloadUrl.ToString(), e.NewVersionInfo));
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
    }
}

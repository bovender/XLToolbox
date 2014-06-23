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
using System.Net;
using XLToolbox.Version;

namespace XLToolbox
{
    /// <summary>
    /// Interaction logic for WindowDownloadUpdate.xaml
    /// </summary>
    public partial class WindowDownloadUpdate : Window
    {
        public Updater Updater { get; private set; }

        public WindowDownloadUpdate(Updater updater, string targetDir)
        {
            InitializeComponent();
            Updater = updater;
            updater.DownloadProgressChanged += updater_DownloadProgressChanged;
            updater.DownloadInstallable += updater_DownloadInstallable;
            updater.DownloadFailedVerification += updater_DownloadFailedVerification;
            updater.DownloadUpdate(targetDir);
        }

        void updater_DownloadFailedVerification(object sender, UpdateAvailableEventArgs e)
        {
            System.Windows.MessageBox.Show(Strings.DownloadedFileCannotBeInstalled,
                Strings.DownloadingXLToolboxUpdate, MessageBoxButton.OK, MessageBoxImage.Warning);
            CloseWindow();
        }

        void updater_DownloadInstallable(object sender, UpdateAvailableEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show(
                String.Format(Strings.UpdateHasBeenDownloaded, Updater.DownloadPath),
                Strings.UpdateAvailable, MessageBoxButton.OKCancel, MessageBoxImage.Information);
            if (result == MessageBoxResult.OK)
            {
                Globals.ThisAddIn.Updater = Updater;
            }
            else
            {
                Globals.ThisAddIn.Updater = null;
            };
            CloseWindow();
        }

        void updater_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            Progress.Value = e.ProgressPercentage;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            // Updater will take care of deleting the partially downloaded file.
            Updater.CancelDownload();
            this.Close();
        }

        /// <summary>
        /// Closes the update progress window. This method invokes the WPF dispatcher
        /// and can be called from event handler that belong to a different thread.
        /// </summary>
        private void CloseWindow()
        {
            this.Dispatcher.Invoke(new Action(Close));
        }
    }
}

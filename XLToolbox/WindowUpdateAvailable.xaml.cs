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
using System.Windows.Forms;
using XLToolbox.Version;

namespace XLToolbox
{
    /// <summary>
    /// Interaction logic for WindowUpdateAvailable.xaml
    /// </summary>
    public partial class WindowUpdateAvailable : Window
    {
        private Updater Updater { get; set; }

        public WindowUpdateAvailable(Updater updater)
        {
            InitializeComponent();
            Updater = updater;
            CurrentVersion.Text = SemanticVersion.CurrentVersion().ToString();
            NewVersion.Text = updater.NewVersion.ToString();
            UpdateDescription.Text = updater.UpdateDescription;
        }

        private void Download_Click(object sender, RoutedEventArgs e)
        {
            string defaultPath = Properties.Settings.Default.DownloadPath;
            if (defaultPath.Length == 0)
            {
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
            };
            FolderBrowserDialog f = new FolderBrowserDialog();
            f.SelectedPath = defaultPath;
            f.ShowNewFolderButton = true;
            f.ShowDialog();
            if (f.SelectedPath.Length > 0)
            {
                Properties.Settings.Default.DownloadPath = f.SelectedPath;
                Properties.Settings.Default.Save();

                WindowDownloadUpdate w = new WindowDownloadUpdate(Updater, f.SelectedPath);
                /* If the update file was previously downloaded, the WindowDownloadUpdate window
                 * will be closed by the ensuing event handlers before we even get a chance to
                 * Show() it. To prevent this race condition, a flag Downloaded was introduced
                 * in the Updater class that signals if the file is downloaded or not. If it is,
                 * we won't Show() the window.
                 */
                if (!Updater.Downloaded)
                {
                    w.Show();
                }
                this.Close();
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}

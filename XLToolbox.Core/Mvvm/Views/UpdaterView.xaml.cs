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
using Bovender.Versioning;

namespace XLToolbox.Mvvm.Views
{
    /// <summary>
    /// Interaction logic for WindowUpdateAvailable.xaml
    /// </summary>
    public partial class UpdaterView : Window
    {
        public UpdaterView(Updater updater)
        {
            InitializeComponent();
        }

        private void Download_Click(object sender, RoutedEventArgs e)
        {

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
    }
}

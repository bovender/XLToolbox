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
using XLToolbox.Versioning;

namespace XLToolbox.Mvvm.Views
{
    /// <summary>
    /// Interaction logic for NotAuthorizedForUpdateView.xaml
    /// </summary>
    public partial class NotAuthorizedForUpdateView : Window
    {
        public Updater Updater { get; private set; }

        public NotAuthorizedForUpdateView(Updater updater)
        {
            InitializeComponent();
            Updater = updater;
            DownloadUrl.NavigateUri = Updater.DownloadUri;
            DownloadUrlLabel.Text = Updater.DownloadUri.ToString();
        }

        private void DownloadUrl_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(Updater.DownloadUri.ToString());
        }
    }
}

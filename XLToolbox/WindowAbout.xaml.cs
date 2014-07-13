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
using XLToolbox.Core;

namespace XLToolbox
{
    /// <summary>
    /// Interaction logic for WindowAbout.xaml
    /// </summary>
    public partial class WindowAbout : Window
    {
        public WindowAbout()
        {
            InitializeComponent();
            TextVersion.Text = String.Format(Strings.VersionParametrized,
                SemanticVersion.CurrentVersion());
        }

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ButtonWebsite_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(Constants.WEBSITE);
        }

        private void ButtonLicense_Click(object sender, RoutedEventArgs e)
        {
            WindowHtmlViewer w = new WindowHtmlViewer(Strings.License, "html/GPLv3.html");
            w.ShowDialog();
        }

        private void ButtonCredits_Click(object sender, RoutedEventArgs e)
        {
            WindowHtmlViewer w = new WindowHtmlViewer(Strings.Credits, "html/credits.html");
            w.ShowDialog();
        }
    }
}

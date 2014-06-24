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
using XLToolbox.Help;

namespace XLToolbox
{
    /// <summary>
    /// Interaction logic for WindowGreeter.xaml
    /// </summary>
    public partial class WindowGreeter : Window
    {
        public WindowGreeter()
        {
            InitializeComponent();
            VersionInfo.Text = String.Format(
                Strings.ThisIsVersion, SemanticVersion.CurrentVersion());
        }

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ButtonDonate_Click(object sender, RoutedEventArgs e)
        {
            Provider.Show(Topic.Donate);
            Close();
        }

        private void ButtonWhatsNew_Click(object sender, RoutedEventArgs e)
        {
            Provider.Show(Topic.WhatsNew);
            Close();
        }
    }
}

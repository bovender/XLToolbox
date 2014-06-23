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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

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
using XLToolbox.Error;

namespace XLToolbox
{
    /// <summary>
    /// Interaction logic for WindowErrorReport.xaml
    /// </summary>
    public partial class WindowErrorReport : Window
    {
        public Reporter Reporter { get; private set; }

        public WindowErrorReport(Reporter r)
        {
            InitializeComponent();
            Reporter = r;
            this.DataContext = r;
        }

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

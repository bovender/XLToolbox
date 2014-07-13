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
using System.Windows.Resources;
using System.IO;

namespace XLToolbox
{
    /// <summary>
    /// Interaction logic for WindowHtmlViewer.xaml
    /// </summary>
    public partial class WindowHtmlViewer : Window
    {
        public WindowHtmlViewer(string title, string docPath)
        {
            InitializeComponent();
            Title = title;
            StreamResourceInfo i = Application.GetResourceStream(new Uri(
                @"pack://application:,,,/XLToolbox;component/" + docPath));
            Browser.NavigateToStream(i.Stream);
        }

        private void ButtonClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}

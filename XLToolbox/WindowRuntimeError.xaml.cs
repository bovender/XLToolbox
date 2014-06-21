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

namespace XLToolbox
{
    /// <summary>
    /// Interaction logic for WindowRuntimeError.xaml
    /// </summary>
    public partial class WindowRuntimeError : Window
    {
        public WindowRuntimeError()
        {
            InitializeComponent();
        }

        public WindowRuntimeError(Exception e) : this()
        {
            ErrorMessage.Text = e.ToString();
            ErrorDescription.Text = e.Message;
        }
    }
}

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
using Bovender.Versioning;

namespace XLToolbox.Versioning
{
    /// <summary>
    /// Interaction logic for UpdaterProcessView.xaml. This view can be
    /// used to display a progress bar while checking for update availability,
    /// or while downloading an update.
    /// </summary>
    public partial class UpdaterProcessView : Window
    {
        public UpdaterProcessView()
        {
            InitializeComponent();
        }
    }
}

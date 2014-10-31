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
using XLToolbox.ExceptionHandler;

namespace XLToolbox.Mvvm.Views
{
    /// <summary>
    /// Interaction logic for ExceptionView.xaml
    /// </summary>
    public partial class ExceptionView : Window
    {
        ExceptionViewModel Reporter { get; set; }
        private object _oldSendButtonContent;
        private Cursor _oldCursor;

        public WindowRuntimeError()
        {
            InitializeComponent();
            _oldSendButtonContent = ButtonSend.Content;
            _oldCursor = Cursor;
        }

        public WindowRuntimeError(ExceptionViewModel r) : this()
        {
            Reporter = r;
            Reporter.UploadSuccessful += Reporter_UploadSuccessful;
            Reporter.UploadFailed += Reporter_UploadFailed;
            this.DataContext = r;
        }

        private void ButtonInfo_Click(object sender, RoutedEventArgs e)
        {
            WindowErrorReport w = new WindowErrorReport(Reporter);
            w.ShowDialog();
        }

        private void ButtonSend_Click(object sender, RoutedEventArgs e)
        {
            ButtonSend.Content = Strings.SendingEllipsis;
            Cursor = Cursors.Wait;
            ButtonSend.IsEnabled = false;
            Reporter.Send();
        }

        void Reporter_UploadFailed(object sender, UploadFailedEventArgs e)
        {
            Cursor = _oldCursor;
            MessageBox.Show(String.Format(Strings.SendingErrorReportFailed, e.Error.ToString()),
                Title, MessageBoxButton.OK, MessageBoxImage.Warning);
            ButtonSend.Content = _oldSendButtonContent;
            ButtonSend.IsEnabled = true;
        }

        void Reporter_UploadSuccessful(object sender, System.Net.UploadValuesCompletedEventArgs e)
        {
            Cursor = _oldCursor;
            MessageBox.Show(Strings.SendingErrorReportSuccessful, Title, MessageBoxButton.OK,
                MessageBoxImage.Information);
            CloseWindow();
        }

        /// <summary>
        /// Closes the window. Uses dispatcher, therefore thread-safe.
        /// </summary>
        private void CloseWindow()
        {
            this.Dispatcher.Invoke(new Action(Close));
        }
    }
}

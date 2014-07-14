using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Threading;

namespace XLToolbox
{
    public static class WpfHelpers
    {
        /// <summary>
        /// Helper method to show modeless WPF window in Excel.
        /// </summary>
        /// <remarks>
        /// This helper is required because a simple Show() results in a modeless
        /// window that does not receive keyboard events.
        /// Cf. http://stackoverflow.com/a/5884085/270712
        /// </remarks>
        /// <typeparam name="T"></typeparam>
        public static void ShowModelessInExcel<T>() where T: Window, new()
        {
            Thread thread = new Thread(() =>
            {
                Window w = new T();
                w.Show();
                w.Closed += (sender2, e2) => w.Dispatcher.InvokeShutdown();
                System.Windows.Threading.Dispatcher.Run();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }
    }
}

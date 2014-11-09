using System;
using System.Windows;
using Bovender.Mvvm.ViewModels;
using System.Threading;

namespace Bovender.Mvvm
{
    public static class Workarounds
    {
        #region Public static methods

        /// <summary>
        /// Helper method to show modeless WPF window in Excel.
        /// </summary>
        /// <remarks>
        /// This helper is required because a simple Show() results in a modeless
        /// window that does not receive keyboard events.
        /// Cf. http://stackoverflow.com/a/5884085/270712
        /// </remarks>
        /// <typeparam name="T"></typeparam>
        public static void ShowModelessInExcel<T>() where T : Window, new()
        {
            Thread thread = new Thread(() =>
            {
                Window view = new T();
                view.Closed += (sender2, e2) => view.Dispatcher.InvokeShutdown();
                view.Show();
                System.Windows.Threading.Dispatcher.Run();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        /// <summary>
        /// Injects a ViewModel into a newly created View and shows it in its own
        /// thread in order to prevent problems with keyboard entry in Excel.
        /// </summary>
        /// <remarks>
        /// This helper is required because a simple Show() results in a modeless
        /// window that does not receive keyboard events.
        /// Cf. http://stackoverflow.com/a/5884085/270712
        /// </remarks>
        /// <typeparam name="T">View class that is derived from <see cref="System.Windows.Window"/>
        /// </typeparam>
        public static void ShowModelessInExcel<T>(ViewModelBase viewModel) where T : Window, new()
        {
            Thread thread = new Thread(() =>
            {
                Window view = viewModel.InjectInto<T>();
                view.Closed += (sender2, args2) => view.Dispatcher.InvokeShutdown();
                view.Show();
                System.Windows.Threading.Dispatcher.Run();
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        #endregion
    }
}

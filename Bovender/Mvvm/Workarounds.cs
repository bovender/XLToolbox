/* Workarounds.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
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

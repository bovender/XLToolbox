using Microsoft.Office.Tools;
/* SheetManagerPane.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2016 Daniel Kraus
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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.SheetManager
{
    /// <summary>
    /// Singleton class that handles the Worksheet Manager task pane.
    /// </summary>
    public class SheetManagerPane
    {
        #region Singleton factory

        public static SheetManagerPane Default
        {
            get
            {
                return _lazy.Value;
            }
        }

        /// <summary>
        /// Gets whether the sheet manager task pane has been initialized and
        /// visible. Other than accessing the Default.Visible property, this
        /// method won't cause the singleton to be instantiated.
        /// </summary>
        public static bool InitializedAndVisible
        {
            get
            {
                if (_lazy.IsValueCreated)
                {
                    return Default.Visible;
                }
                else
                {
                    return false;
                }
            }
        }

        #endregion

        #region Public properties

        public bool Visible
        {
            get
            {
                return _visible;
            }
            set
            {
                Logger.Info("Visible: Set: {0}", value);
                _visible = value;
                foreach (CustomTaskPane pane in Panes.Values)
                {
                    pane.Visible = value;
                }
            }
        }

        public int Width
        {
            get
            {
                return _width;
            }
            set
            {
                Logger.Info("Width: Set: {0}", value);
                _width = value;
                foreach (CustomTaskPane pane in Panes.Values)
                {
                    pane.Width = value;
                }
            }
        }

        #endregion

        #region Events

        /// <summary>
        /// Raised when the sheet manager singleton has been initialized.
        /// Caveat: At the time when the event is raised, the static Default
        /// property will not yet return the instance. Subscribers to this
        /// event should use the SheetManagerEventArgs.Instance property
        /// to access the singleton instance.
        /// </summary>
        /// <remarks>
        /// This is a static event. Subscribers should take care to unsubscribe
        /// from it, otherwise they will never be garbage-collected.
        /// </remarks>
        public static event EventHandler<SheetManagerEventArgs> SheetManagerInitialized;

        /// <summary>
        /// Raised when the visibility of the encapsulated task pane changed.
        /// </summary>
        public event EventHandler<SheetManagerEventArgs> VisibilityChanged;

        #endregion

        #region Public methods

        /// <summary>
        /// Updates the SheetManagerPane for the active window.
        /// This method should be called from a WindowActivate
        /// event handler.
        /// </summary>
        public void UpdatePanes()
        {

        }
        
        #endregion

        #region Constructor

        private SheetManagerPane()
        {
            _width = UserSettings.UserSettings.Default.TaskPaneWidth;
            _viewModel = new WorkbookViewModel(Instance.Default.ActiveWorkbook);
            AttachToCurrentWindow();
            Excel.ViewModels.Instance.Default.Application.WindowActivate += Application_WindowActivate;
        }

        #endregion

        #region Private methods

        private static void OnInitialized(SheetManagerPane sheetManagerPane)
        {
            EventHandler<SheetManagerEventArgs> h = SheetManagerInitialized;
            if (h != null)
            {
                h(null, new SheetManagerEventArgs(sheetManagerPane));
            }
        }

        private void OnVisibilityChanged(CustomTaskPane senderTaskPane)
        {
            if (!_lockVisibleChangeEventHandler)
            {
                _lockVisibleChangeEventHandler = true;
                Logger.Info("OnVisibilityChanged");

                // Synchronize the visibility of all task panes.
                // We cannot use our own Visible property to accomplish this,
                // because accessing the property of the task pane that raised
                // the event causes an exception.
                _visible = senderTaskPane.Visible;
                foreach (CustomTaskPane p in Panes.Values)
                {
                    if (senderTaskPane != p)
                    {
                        p.Visible = _visible;
                    }
                }

                UserSettings.UserSettings.Default.SheetManagerVisible = Visible;
                if (Visible)
                {
                    _viewModel.MonitorWorkbook.Execute(null);
                }
                else
                {
                    _viewModel.UnmonitorWorkbook.Execute(null);
                }
                EventHandler<SheetManagerEventArgs> h = VisibilityChanged;
                if (h != null)
                {
                    h(this, new SheetManagerEventArgs(this));
                }
                _lockVisibleChangeEventHandler = false;
            }
        }

        private void Application_WindowActivate(
            Microsoft.Office.Interop.Excel.Workbook Wb,
            Microsoft.Office.Interop.Excel.Window Wn)
        {
            Logger.Info("Application_WindowActivate");
            AttachToCurrentWindow();
        }

        private void AttachToCurrentWindow()
        {
            // If the current window does not yet have our task pane, add it to it
            IntPtr currentHandle = Bovender.Win32Window.MainWindowHandleProvider();
            if (!Panes.ContainsKey(currentHandle))
            {
                Logger.Info("Attaching new WorksheetManager panel to window 0x{0:X08}", currentHandle);
                UserControl userControl = new UserControl();
                SheetManagerControl view = new SheetManagerControl() { DataContext = _viewModel };
                ElementHost elementHost = new ElementHost() { Child = view };
                userControl.Controls.Add(elementHost);
                elementHost.Dock = DockStyle.Fill;
                CustomTaskPane pane = Globals.CustomTaskPanes.Add(userControl, Strings.WorksheetManager);
                Panes.Add(currentHandle, pane);
                pane.Width = Width;
                pane.Visible = Visible;
                pane.VisibleChanged += (sender, args) =>
                {
                    OnVisibilityChanged(sender as CustomTaskPane);
                };
            }
            else
            {
                Logger.Info("Window 0x{0:X08} already has a WorksheetManager panel", currentHandle);
            }
        }

        #endregion

        #region Private properties

        /// <summary>
        /// Manages the SheetManager task panes for individual Excel windows.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Excel 2013 is an SDI application, which means it has multiple windows
        /// for multiple workbooks. Excel 2010 is an MDI application, which means
        /// that multiple open workbooks were shown in just a single application
        /// window. The new SDI mode has consequences for task panes, which are
        /// bound to each window.
        /// </para>
        /// <para>
        /// Inspired by an answer by @antonio-nakic-alfirevic on StackOverflow:
        /// http://stackoverflow.com/a/24732000/270712
        /// </para>
        /// <para>
        /// More at https://msdn.microsoft.com/en-us/library/office/dn251093.aspx
        /// </para>
        /// </remarks>
        private Dictionary<IntPtr, CustomTaskPane> Panes
        {
            get
            {
                return _lazyPanes.Value;
            }
        }

        #endregion

        #region Private fields

        private WorkbookViewModel _viewModel;
        private bool _visible;
        private int _width;
        private bool _lockVisibleChangeEventHandler;

        #endregion

        #region Private static fields

        private static readonly Lazy<SheetManagerPane> _lazy = new Lazy<SheetManagerPane>(
            () =>
            {
                Logger.Info("Lazily creating SheetManagerPane instance");
                SheetManagerPane p = new SheetManagerPane();
                OnInitialized(p);
                return p;
            }
        );

        private static readonly Lazy<Dictionary<IntPtr, CustomTaskPane>> _lazyPanes =
            new Lazy<Dictionary<IntPtr, CustomTaskPane>>(() =>
            {
                return new Dictionary<IntPtr, CustomTaskPane>();
            });

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

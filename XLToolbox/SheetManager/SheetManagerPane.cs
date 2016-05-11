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
                return _pane.Visible;
            }
            set
            {
                _pane.Visible = value;
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

        #region Constructor

        private SheetManagerPane()
        {
            CreateTaskPane();
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

        private void OnVisibilityChanged()
        {
            UserSettings.Default.SheetManagerVisible = _pane.Visible;
            if (_pane.Visible)
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
        }

        private void CreateTaskPane()
        {
            _viewModel = new WorkbookViewModel(Instance.Default.ActiveWorkbook);
            UserControl userControl = new UserControl();
            SheetManagerControl view = new SheetManagerControl() { DataContext = _viewModel };
            ElementHost elementHost = new ElementHost() { Child = view };
            userControl.Controls.Add(elementHost);
            elementHost.Dock = DockStyle.Fill;
            _pane = Globals.CustomTaskPanes.Add(userControl, Strings.WorksheetManager);
            _pane.Width = UserSettings.Default.TaskPaneWidth;
            _pane.VisibleChanged += (sender, args) =>
            {
                OnVisibilityChanged();
            };
        }

        #endregion

        #region Private fields

        private WorkbookViewModel _viewModel;
        private CustomTaskPane _pane;

        #endregion

        #region Private static fields

        private static Lazy<SheetManagerPane> _lazy = new Lazy<SheetManagerPane>(
            () =>
            {
                SheetManagerPane p = new SheetManagerPane();
                OnInitialized(p);
                return p;
            }
        );

        #endregion
    }
}

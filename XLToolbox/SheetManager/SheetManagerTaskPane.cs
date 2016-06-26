using Microsoft.Office.Tools;
/* SheetManagerTaskPane.cs
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
    public class SheetManagerTaskPane
    {
        #region Public properties

        public WorkbookViewModel ViewModel { get; private set; }

        public int Width
        {
            get
            {
                if (_pane != null)
                {
                    return _pane.Width;
                }
                else
                {
                    throw new InvalidOperationException("Custom task pane not initialized");
                }
            }
            set
            {
                if (_pane != null)
                {
                    _pane.Width = value;
                    _width = value;
                }
                else
                {
                    throw new InvalidOperationException("Custom task pane not initialized");
                }
            }
        }

        public bool Visible
        {
            get
            {
                if (_pane != null)
                {
                    return _pane.Visible;
                }
                else
                {
                    throw new InvalidOperationException("Custom task pane not initialized");
                }
            }
            set
            {
                if (_pane != null)
                {
                    if (value != _visible)
                    {
                        _pane.Visible = value;
                        _visible = value;
                        OnVisibilityChanged();
                    }
                }
                else
                {
                    throw new InvalidOperationException("Custom task pane not initialized");
                }
            }
        }

        #endregion

        #region Events

        public event EventHandler<SheetManagerEventArgs> VisibilityChanged;

        #endregion

        #region Constructor

        public SheetManagerTaskPane(WorkbookViewModel workbookViewModel, int initialWidth, bool visible)
        {
            if (workbookViewModel == null)
            {
                throw new ArgumentNullException("workbookViewModel", "WorkbokViewModel must not be null");
            }
            ViewModel = workbookViewModel;
            _width = initialWidth;
            _visible = visible;
            CreateCustomTaskPane();
        }

        #endregion

        #region Protected methods

        protected void MonitorWorkbook(bool visible)
        {
            if (visible)
            {
                ViewModel.MonitorWorkbook.Execute(null);
            }
            else
            {
                ViewModel.UnmonitorWorkbook.Execute(null);
            }
        }

        protected virtual void OnVisibilityChanged()
        {
            MonitorWorkbook(_visible);
            EventHandler<SheetManagerEventArgs> h = VisibilityChanged;
            if (h != null)
            {
                h(this, new SheetManagerEventArgs(this));
            }
        }

        #endregion

        #region Private methods

        private void CreateCustomTaskPane()
        {
            Logger.Info("CreateCustomTaskPane");
            UserControl userControl = new UserControl();
            SheetManagerControl view = new SheetManagerControl() { DataContext = ViewModel };
            ElementHost elementHost = new ElementHost() { Child = view };
            userControl.Controls.Add(elementHost);
            elementHost.Dock = DockStyle.Fill;
            _pane = Globals.CustomTaskPanes.Add(userControl, Strings.WorksheetManager);
            _pane.Width = _width;
            _pane.Visible = _visible;
            MonitorWorkbook(_visible);
            _pane.VisibleChanged += (sender, args) =>
            {
                _visible = _pane.Visible;
                OnVisibilityChanged();
            };

        }

        #endregion

        #region Private fields

        private CustomTaskPane _pane;
        private bool _visible;
        private int _width;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

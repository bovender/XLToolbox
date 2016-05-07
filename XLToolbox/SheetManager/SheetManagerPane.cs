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
    internal class SheetManagerPane
    {
        #region Singleton factory

        public static SheetManagerPane Default
        {
            get
            {
                return _lazy.Value;
            }
        }

        #endregion

        #region Public methods

        public void Show()
        {
            _pane.Visible = true;
        }

        public void Hide()
        {
            _pane.Visible = false;
        }

        #endregion

        #region Constructor

        private SheetManagerPane()
        {
            CreateTaskPane();
        }

        #endregion

        #region Private methods

        private void CreateTaskPane()
        {
            WorkbookViewModel wvm = new WorkbookViewModel(Instance.Default.ActiveWorkbook);
            UserControl userControl = new UserControl();
            SheetManagerControl view = new SheetManagerControl() { DataContext = wvm };
            ElementHost elementHost = new ElementHost() { Child = view };
            userControl.Controls.Add(elementHost);
            elementHost.Dock = DockStyle.Fill;
            _pane = Globals.CustomTaskPanes.Add(userControl, Strings.WorksheetManager);
            _pane.Width = UserSettings.Default.TaskPaneWidth;
        }

        #endregion

        #region Private fields

        private CustomTaskPane _pane;

        #endregion

        #region Private static fields

        private static Lazy<SheetManagerPane> _lazy = new Lazy<SheetManagerPane>(
            () =>
            {
                return new SheetManagerPane();
            }
        );

        #endregion
    }
}

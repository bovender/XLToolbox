/* Manager.cs
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
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Keyboard
{
    /// <summary>
    /// Manages keyboard shortcuts.
    /// </summary>
    public class Manager
    {
        private const string ADDIN_RESOURCE_NAME = "XLToolbox.Keyboard.XLToolboxKeyboardBridge.xlam";

        #region Singleton factory

        public static Manager Default
        {
            get { return _lazy.Value; }
        }
        
        #endregion

        #region Public properties

        public IList<Shortcut> Shortcuts { get; private set; }

        #endregion

        #region Public methods

        public void EnableShortcuts()
        {
            foreach (Shortcut shortcut in Shortcuts)
            {
                shortcut.Enable();
            }
        }

        public void DisableShortcuts()
        {
            foreach (Shortcut shortcut in Shortcuts)
            {
                shortcut.Disable();
            }
        }

        /// <summary>
        /// Resets the shortcut collection to built-in defaults.
        /// </summary>
        public void SetDefaults()
        {
            Shortcuts = new List<Shortcut>();
            Shortcuts.Add(new Shortcut("^+%Q", Command.QuitExcel)); // CTRL SHIFT ALT Q
        }

        #endregion

        #region Constructor

        private Manager()
        {
            SetDefaults();
        }

        #endregion

        #region Private static fields

        private static Lazy<Manager> _lazy = new Lazy<Manager>(
            () =>
            {
                Instance.Default.LoadAddinFromEmbeddedResource(ADDIN_RESOURCE_NAME);
                return new Manager();
            }
        );

        #endregion
    }
}

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
using System.Collections.ObjectModel;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using XLToolbox.Excel.ViewModels;
using Bovender.Extensions;

namespace XLToolbox.Keyboard
{
    /// <summary>
    /// Manages keyboard shortcuts.
    /// </summary>
    public class Manager : IDisposable
    {
        private const string ADDIN_FILENAME = "XLToolboxKeyboardBridge.xlam";
        private const string ADDIN_RESOURCE_NAME = "XLToolbox.Keyboard." + ADDIN_FILENAME;

        #region Singleton factory

        public static Manager Default
        {
            get { return _lazy.Value; }
        }
        
        #endregion

        #region Public properties

        public bool IsEnabled
        {
            get
            {
                return _isEnabled;
            }
            set
            {
                if (value != _isEnabled)
                {
                    if (value)
                    {
                        RegisterShortcuts();
                    }
                    else
                    {
                        UnregisterShortcuts();
                    }
                }
                _isEnabled = value;
            }
        }

        public ObservableCollection<Shortcut> Shortcuts
        {
            get
            {
                return _shortcuts;
            }
            set
            {
                MergeSubset(value);
            }
        }

        #endregion

        #region Public methods

        public void RegisterShortcuts()
        {
            if (IsEnabled && !_registered)
            {
                foreach (Shortcut shortcut in Shortcuts)
                {
                    shortcut.Register();
                }
                _registered = true;
            }
        }

        public void UnregisterShortcuts()
        {
            if (_registered)
            {
                foreach (Shortcut shortcut in Shortcuts)
                {
                    shortcut.Unregister();
                }
                _registered = false;
            }
        }

        public void SetShortcut(Command command, string keySequence)
        {
            Shortcut shortcut = Shortcuts.First(s => s.Command == command);
            shortcut.KeySequence = keySequence;
            shortcut.Register();
        }

        public void UnsetShortcut(Command command)
        {
            Shortcut shortcut = Shortcuts.First(s => s.Command == command);
            shortcut.Unregister();
            shortcut.KeySequence = String.Empty;
        }

        /// <summary>
        /// Resets the shortcut collection to built-in defaults.
        /// </summary>
        public void SetDefaults()
        {
            UnregisterShortcuts();
            CreateListOfCommands();
            SetShortcut(Command.QuitExcel, "^+%Q");
            SetShortcut(Command.SaveAs, "^+S");
            SetShortcut(Command.AnovaRepeat, "^+N");
            SetShortcut(Command.LastErrorBars, "^+E");
            SetShortcut(Command.ChartDesign, "^+D");
            SetShortcut(Command.FormulaBuilder, "^+B");
            SetShortcut(Command.MoveDataSeriesLeft, "^+{LEFT}");
            SetShortcut(Command.MoveDataSeriesRight, "^+{RIGHT}");
            SetShortcut(Command.SelectAllShapes, "^+A");
            RegisterShortcuts();
        }

        #endregion

        #region Constructor

        private Manager()
        {
            _isEnabled = true;
            SetDefaults();
        }

        #endregion

        #region Disposing
        
        ~Manager()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool calledFromPublicMethod)
        {
            if (!_disposed)
            {
                _disposed = true;
                if (calledFromPublicMethod)
                {
                    Instance.Default.Application.Workbooks[ADDIN_FILENAME].Close(SaveChanges: false);
                }
                try
                {
                    System.IO.File.Delete(_tempFile);
                }
                catch (Exception)
                {
                    // TODO: Log errors
                }
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Creates a new list of Shortcuts with one shortcut for each
        /// XL Toolbox command.
        /// </summary>
        private void CreateListOfCommands()
        {
            _shortcuts = new ObservableCollection<Shortcut>();
            IEnumerable<Command> commands = ((Command[])Enum.GetValues(typeof(Command))).OrderBy(c => c.ToString());
            foreach (Command command in commands)
            {
                _shortcuts.Add(new Shortcut(String.Empty, command));
            }
        }

        /// <summary>
        /// Merges a subset of shortcuts, e.g. deserialized from XLToolbox.UserSettings,
        /// into the list that contains one shortcut for each command.
        /// </summary>
        private void MergeSubset(IList<Shortcut> subset)
        {
            UnregisterShortcuts();
            CreateListOfCommands();
            foreach (Shortcut shortcutInSubset in subset)
            {
                // Since the _shortcuts list contains shortcuts for all values of the
                // Command enum, it should be save to use _shortcuts.First without
                // further checks.
                _shortcuts.First(s => s.Command == shortcutInSubset.Command).KeySequence = shortcutInSubset.KeySequence;
            }
            RegisterShortcuts();
        }
        
        #endregion

        #region Private fields

        private ObservableCollection<Shortcut> _shortcuts;
        private bool _isEnabled;
        private bool _registered;
        private bool _disposed;

        #endregion

        #region Private static fields

        private static Lazy<Manager> _lazy = new Lazy<Manager>(
            () =>
            {
                _tempFile = Instance.Default.LoadAddinFromEmbeddedResource(ADDIN_RESOURCE_NAME);
                return new Manager();
            }
        );

        private static string _tempFile;

        #endregion
    }
}

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
    public class Manager
    {
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
                Logger.Info("RegisterShortcuts");
                foreach (Shortcut shortcut in Shortcuts)
                {
                    shortcut.Register();
                }
                _registered = true;
            }
            else
            {
                Logger.Info("RegisterShortcuts: not enabled or registered already");
            }
        }

        public void UnregisterShortcuts()
        {
            if (_registered)
            {
                Logger.Info("UnregisterShortcuts");
                foreach (Shortcut shortcut in Shortcuts)
                {
                    shortcut.Unregister();
                }
                _registered = false;
            }
            else
            {
                Logger.Info("UnregisterShortcuts: not registered");
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
            Logger.Info("SetDefaults");
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
                Legacy.LegacyToolbox l = Legacy.LegacyToolbox.Default;
                return new Manager();
            }
        );

        private static string _tempFile;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

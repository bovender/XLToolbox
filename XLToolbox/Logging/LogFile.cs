using NLog;
using NLog.Config;
using NLog.Targets;
using NLog.Targets.Wrappers;
/* LogFile.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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

namespace XLToolbox.Logging
{
    /// <summary>
    /// Provides logging to file and to the debug console; wraps
    /// NLog configuration and targets.
    /// </summary>
    public class LogFile : Bovender.Logging.LogFile
    {
        #region Singleton

        new public static LogFile Default { get { return _lazy.Value; } }

        private static readonly Lazy<LogFile> _lazy = new Lazy<LogFile>(
            () => 
            {
                LogFile logFile = new LogFile();
                Bovender.Logging.LogFile.LogFileProvider = new Func<Bovender.Logging.LogFile>(() => logFile);
                return logFile;
            });

        #endregion

        #region Static properties

        /// <summary>
        /// Gets whether file logging is enabled, without initializing
        /// the singleton instance if it isn't.
        /// </summary>
        new public static bool IsInitializedAndEnabled
        {
            get
            {
                return _lazy.IsValueCreated && Default.IsFileLoggingEnabled;
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the folder where log files are stored.
        /// </summary>
        public override string LogFolder
        {
            get
            {
                if (_logFolder == null)
                {
                    _logFolder = System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        Properties.Settings.Default.AppDataFolder,
                        Properties.Settings.Default.UserFolder);
                }
                return _logFolder;
            }
        }

        #endregion

        #region Constructor

        private LogFile()
            : base()
        { }

        #endregion

        #region Private fields

        string _logFolder;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

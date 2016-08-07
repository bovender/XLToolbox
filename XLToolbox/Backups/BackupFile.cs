/* BackupFile.cs
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
using System.IO;

namespace XLToolbox.Backups
{
    /// <summary>
    /// Represents a backup file with a date stamp in the file name.
    /// </summary>
    class BackupFile
    {
        #region Properties

        public DateTime TimeStamp { get; private set; }

        #endregion

        #region Public methods

        /// <summary>
        /// Deletes the physical file that is represented by the
        /// BackupFile object.
        /// </summary>
        /// <returns>True if the file was deleted, false if not.</returns>
        public bool Delete()
        {
            bool result = false;
            try 
	        {	        
                File.Delete(_path);
                result = true;
	        }
	        catch { }
            return result;
        }
        
        #endregion

        #region Constructor

        /// <summary>
        /// Creates a new BackupFile object from a given backup file path.
        /// </summary>
        /// <param name="path"></param>
        public BackupFile(string path)
        {
            _path = path;
        }

        #endregion

        #region Private methods

        private void ParseTimeStamp(string path)
        {
            string dir = Path.GetDirectoryName(path);
            string fn = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);
        }

        #endregion

        #region Fields

        private string _path;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

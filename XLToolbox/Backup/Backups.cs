/* BackupFiles.cs
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

namespace XLToolbox.Backup
{
    /// <summary>
    /// Represents a collection of backups for a given file.
    /// </summary>
    public class Backups
    {
        #region Properties

        /// <summary>
        /// Gets the entire path to the file being backed up.
        /// </summary>
        public string FilePath { get; private set; }

        /// <summary>
        /// Gets or sets the relative directory where backups are
        /// stored.
        /// </summary>
        public string BackupDir
        {
            get
            {
                return _backupDir;
            }
            set
            {
                if (!Path.IsPathRooted(value))
                {
                    _backupDir = value;
                    _backupPath = null;
                }
                else
	            {
                    throw new ArgumentException("The backup directory must be a relative directory", BackupDir);
	            }
            }
        }

        /// <summary>
        /// Gets the full backup path for the file.
        /// </summary>
        public string BackupPath
        {
            get
            {
                if (String.IsNullOrEmpty(_backupPath))
                {
                    _backupPath = Path.Combine(
                        Path.GetDirectoryName(FilePath),
                        BackupDir);
                }
                return _backupPath;
            }
        } 

        public List<BackupFile> Files { get; private set; }

        public int Count
        {
            get
            {
                if (Files != null)
                {
                    return Files.Count;
                }
                else
                {
                    return 0;
                }
            }
        }

        #endregion

        #region Public methods

        public void BuildList()
        {
            try
            {
                IEnumerable<string> files = Directory.EnumerateFiles(
                    BackupPath,
                    String.Format("*{0}*", TimeStamp.WildcardPattern),
                    SearchOption.TopDirectoryOnly);
                Files = files.Select(f => new BackupFile(f)).Where(f => f.IsValidBackup).ToList();
            }
            catch
            {
                Files = null;
            }
        }

        public bool Create()
        {
            bool result = false;
            try
            {
                DateTime dt = File.GetLastWriteTime(FilePath);
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Purges the backup files: Keeps all files of today, one for
        /// the last week, one for the month before last week, and one
        /// for all past years.
        /// </summary>
        public void Purge()
        {

        }

        /// <summary>
        /// Removes all backup files.
        /// </summary>
        /// <returns>True if all files were delete, false if at least
        /// one backup file could not be deleted.</returns>
        public bool DeleteAllBackups()
        {
            bool result = true;
            foreach (BackupFile file in Files)
            {
                result &= file.Delete();
            }
            return result;
        }
        
        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new BackupFiles object for a given file, using the
        /// backup directory from the UserSettings.
        /// </summary>
        public Backups(string filePath)
            : this(filePath, UserSettings.UserSettings.Default.BackupDir)
        { }

        /// <summary>
        /// Creates a new BackupFiles object for a given file and a given
        /// backup directory.
        /// </summary>
        public Backups(string filePath, string backupDir)
        {
            FilePath = filePath;
            BackupDir = backupDir;
            BuildList();
        }
        
        #endregion

        #region Private methods

        #endregion

        #region Fields

        private string _backupDir;
        private string _backupPath;

        #endregion
    }
}

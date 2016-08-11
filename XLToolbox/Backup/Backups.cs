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
        #region Static methods

        public static void Enable()
        {
            if (!IsEnabled)
            {
                _isEnabled = true;
                XLToolbox.Excel.ViewModels.Instance.Default.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            }
        }

        public static void Disable()
        {
            if (IsEnabled)
            {
                _isEnabled = false;
                XLToolbox.Excel.ViewModels.Instance.Default.Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
            }
        }

        protected static void Application_WorkbookBeforeSave(
            Microsoft.Office.Interop.Excel.Workbook Wb,
            bool SaveAsUI, ref bool Cancel)
        {
            string fn = Wb.FullName;
            if (System.IO.File.Exists(fn))
            {
                Logger.Info("Application_WorkbookBeforeSave: Creating backup copy before saving the workbook");
                string dir = UserSettings.UserSettings.Default.BackupDir;
                if (String.IsNullOrEmpty(dir)) dir = ".backup";
                Backups b = new Backups(Wb.FullName, dir);
                b.Create();
            }
            else
            {
                Logger.Info("Application_WorkbookBeforeSave: Skipping backup because file does not exist");
            }
        }

        public static bool IsEnabled
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
                        Enable();
                    }
                    else
                    {
                        Disable();
                    }
                }
            }
        }

        private static bool _isEnabled;

        #endregion

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
                    SearchOption.AllDirectories);
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
            BackupFile bf = BackupFile.CreateBackup(FilePath, BackupDir);
            if (bf != null)
            {
                Logger.Info("Create: Created new backup");
                if (Files != null)
                {
                    Files.Add(bf);
                }
                else
                {
                    Files = new List<BackupFile>() { bf };
                }
                Purge();
                result = true;
            }
            else
            {
                Logger.Warn("Create: Failed to create backup");
            }
            return result;
        }

        /// <summary>
        /// Purges the backup files: Keeps all files of today, one for
        /// the last week, one for the month before last week, and one
        /// for all past years.
        /// </summary>
        public void Purge()
        {
            // Sort files by time stamp reverse
            Files.Sort((a, b) => b.TimeStamp.DateTime.CompareTo(a.TimeStamp.DateTime));
            var enumerator = Files.GetEnumerator();
            enumerator.MoveNext();

            // Keep all backups with today's time stamp.
            Logger.Info("Purge: Skipping today's files.");
            while (enumerator.Current != null && enumerator.Current.IsOfToday)
            {
                enumerator.MoveNext();
            }

            // Keep one per day, totalling 7 days
            DateTime currentDate = DateTime.MinValue;
            Logger.Info("Puge: Keeping one file per day, totalling 7 files");
            int count = 0;
            while (enumerator.Current != null && count < 7)
            {
                DateTime d = enumerator.Current.TimeStamp.DateTime.Date;
                if (d != currentDate)
                {
                    currentDate = d;
                    count++;
                }
                else
                {
                    enumerator.Current.Delete();
                }
                enumerator.MoveNext();
            }

            // Keep one per month, totalling 12 backups
            // Set the current date to the 1st day of the month
            Logger.Info("Puge: Keeping one file per month, totalling 12 files");
            currentDate = currentDate.AddDays(1 - currentDate.Day);
            count = 0;
            while (enumerator.Current != null && count < 12)
            {
                DateTime d = enumerator.Current.TimeStamp.DateTime.Date;
                d = d.AddDays(1 - d.Day);
                if (d != currentDate)
                {
                    currentDate = d;
                    count++;
                }
                else
                {
                    enumerator.Current.Delete();
                }
                enumerator.MoveNext();
            }

            Files.RemoveAll(f => f.IsDeleted);
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

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

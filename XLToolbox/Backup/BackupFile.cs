/* BackupFile.cs
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
using Bovender.Extensions;
using System.Text;
using IO = System.IO;

namespace XLToolbox.Backup
{
    /// <summary>
    /// Represents a backup file with a date stamp in the file name.
    /// </summary>
    public class BackupFile
    {
        #region Factory

        /// <summary>
        /// Creates a new backup by copying the file to backup to the backup dir
        /// and inserting a time stamp with the last write time into the file name.
        /// </summary>
        /// <param name="fileToBackUp">File to back up.</param>
        /// <param name="backupDir">Backup directory; this must not be rooted.</param>
        /// <returns>BackupFile object representing the newly created backup file,
        /// or null if an I/O error occurred.</returns>
        public static BackupFile CreateBackup(string fileToBackUp, string backupDir)
        {
            if (IO.Path.IsPathRooted(backupDir))
	        {
                throw new ArgumentException("Backup dir must be relative", "backupDir");
	        }
            BackupFile bf = null;
            try
            {
                DateTime dt = IO.File.GetLastWriteTime(fileToBackUp);
                string fn =IO.Path.GetFileNameWithoutExtension(fileToBackUp) +
                    dt.ToString(TimeStamp.FormatPattern) +
                    IO.Path.GetExtension(fileToBackUp);
                Logger.Info("CreateBackup: Copying file");
                string dir = IO.Path.Combine(
                    IO.Path.GetDirectoryName(fileToBackUp),
                    backupDir);
                IO.Directory.CreateDirectory(dir);
                string target = IO.Path.Combine(dir, fn);
                IO.File.Copy(fileToBackUp, target);
                Logger.Info("CreateBackup: Copy complete");
                bf = new BackupFile(target);
            }
            catch (IO.IOException e)
            { 
                Logger.Warn("CreateBackup: Failed to create backup");
                Logger.Warn(e);
                bf = new BackupFile(e);
            }
            return bf;
        }

        #endregion

        #region Properties

        public bool IsDeleted { get; private set; }

        /// <summary>
        /// Gets or sets the path of the backup file.
        /// </summary>
        public string Path
        {
            get
            {
                return _path;
            }
            set
            {
                _path = value;
                TimeStamp = new TimeStamp(_path);
            }
        }

        /// <summary>
        /// Gets the TimeStamp object that contains the date and time
        /// extracted from the Path. If the Path does not contain a
        /// valid time stamp, the TimeStamp's HasValue property is false.
        /// </summary>
        public TimeStamp TimeStamp { get; private set; }

        public int Year
        {
            get
            {
                return (TimeStamp != null) ? TimeStamp.DateTime.Year : 0;
            }
        }

        public int Month
        {
            get
            {
                return (TimeStamp != null) ? TimeStamp.DateTime.Month : 0;
            }
        }

        public int Day
        {
            get
            {
                return (TimeStamp != null) ? TimeStamp.DateTime.Day : 0;
            }
        }

        /// <summary>
        /// Gets whether the file's time stamp equals today's date.
        /// If the file does not have a valid time stamp, it returns
        /// false.
        /// </summary>
        public bool IsOfToday
        {
            get
            {
                return (TimeStamp != null) ? TimeStamp.DateTime.Date == DateTime.Today : false;
            }
        }

        public bool IsValidBackup
        {
            get
            {
                return (TimeStamp != null) ?  TimeStamp.HasValue : false;
            }
        }

        public Exception Exception { get; private set; }

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
                System.IO.File.Delete(_path);
                result = true;
	        }
	        catch { }
            IsDeleted = result;
            return result;
        }

        /// <summary>
        /// Opens the workbook represented by this BackupFile.
        /// </summary>
        public void Open()
        {
            Microsoft.Office.Interop.Excel.Workbooks w = XLToolbox.Excel.ViewModels.Instance.Default.Application.Workbooks;
            try
            {
                w.Open(Path, ReadOnly: true);
            }
            catch (Exception e)
            {
                Logger.Warn("Open: Failed to open workbook");
                Logger.Warn(e);
            }
            finally
            {
                Bovender.ComHelpers.ReleaseComObject(w);
            }
        }

        #endregion

        #region Constructor

        protected BackupFile() { }

        /// <summary>
        /// Creates a new BackupFile object from a given backup file path.
        /// </summary>
        /// <param name="path"></param>
        public BackupFile(string path)
            : this()
        {
            Path = path;
        }

        protected BackupFile(Exception exception)
            : this()
        {
            Exception = exception;
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

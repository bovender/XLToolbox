/* Options.cs
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
using Bovender.UserSettings;
using Bovender.Versioning;
using XLToolbox.Export.Models;
using System.Collections.ObjectModel;

namespace XLToolbox
{
    public class UserSettings : UserSettingsBase
    {
        #region Static property

        public static string UserSettingsFile
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    Properties.Settings.Default.AppDataFolder,
                    Properties.Settings.Default.UserSettingsFile);
            }
        }

        #endregion

        #region Singleton factory

        public static UserSettings Default
        {
            get
            {
                return _lazy.Value;
            }
        }

        private static Lazy<UserSettings> _lazy = new Lazy<UserSettings>(() =>
        {
            return FromFileOrDefault<UserSettings>(UserSettingsFile);
        });

        #endregion

        #region User settings

        public string DownloadPath
        {
            get
            {
                return _downloadPath != null ? _downloadPath
                    : Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            }
            set
            {
                _downloadPath = value;
            }
        }

        public string ExportPath
        {
            get
            {
                return _exportPath;
            }
            set
            {
                _exportPath = value;
            }
        }

        public Export.Models.Unit LastExportUnit
        {
            get
            {
                return _lastExportUnit;
            }
            set
            {
                _lastExportUnit = value;
            }
        }

        public int LastExportPreset { get; set; }

        public Export.Models.BatchExportSettings BatchExportSettings
        {
            get
            {
                return _batchExportSettings;
            }
            set
            {
                _batchExportSettings = value;
            }
        }

        public bool WorksheetManagerAlwaysOnTop { get; set; }

        public Csv.CsvFile CsvImport
        {
            get
            {
                if (_csvImport == null)
                {
                    _csvImport = _csvExport;
                }
                return _csvImport;
            }
            set
            {
                _csvImport = value;
            }
        }
        
        public Csv.CsvFile CsvExport
        {
            get
            {
                if (_csvExport == null)
                {
                    _csvExport = _csvImport;
                }
                return _csvExport;
            }
            set
            {
                _csvExport = value;
            }
        }

        public int LastAnova { get; set; }

        public int LastErrorBars { get; set; }

        public int TaskPaneWidth
        {
            get
            {
                if (_taskPaneWidth == 0)
                {
                    _taskPaneWidth = 320;
                }
                return _taskPaneWidth;
            }
            set
            {
                _taskPaneWidth = value;
            }
        }

        public ObservableCollection<Preset> ExportPresets { get; set; }

        #endregion

        #region Auxiliary public methods

        public Preset RetrieveExportPresetOrDefault(int presetIndex)
        {
            if (ExportPresets != null && presetIndex < ExportPresets.Count)
            {
                return ExportPresets[presetIndex];
            }
            else
            {
                return new Preset();
            }
        }

        public Preset RetrieveExportPresetOrDefault()
        {
            return RetrieveExportPresetOrDefault(LastExportPreset);
        }

        public void StoreExportPreset(Preset preset)
        {
            LastExportPreset = GetExportPresetIndex(preset);
        }

        public int GetExportPresetIndex(Preset preset)
        {
            if (ExportPresets != null)
            {
                return ExportPresets.IndexOf(preset);
            }
            else
            {
                return 0;
            }
        }

        #endregion

        #region Implementation of UserSettingsBase

        public override void Save()
        {
            Save(UserSettingsFile);
        }

        protected override void WriteYamlHeader(StreamWriter streamWriter)
        {
            streamWriter.WriteLine(
                String.Format("# {0} <{1}>",
                    Properties.Settings.Default.AddinName,
                    Properties.Settings.Default.WebsiteUrl));
            streamWriter.WriteLine("# User settings file generated by version "
                + XLToolbox.Versioning.SemanticVersion.CurrentVersion().ToString());
            streamWriter.WriteLine("# " + System.DateTime.Now.ToString());
            base.WriteYamlHeader(streamWriter);
        }

        #endregion

        #region Private fields

        private string _downloadPath;
        private string _exportPath;
        private Export.Models.Unit _lastExportUnit;
        private Export.Models.BatchExportSettings _batchExportSettings;
        private Csv.CsvFile _csvImport;
        private Csv.CsvFile _csvExport;
        private int _taskPaneWidth;

        #endregion
    }
}

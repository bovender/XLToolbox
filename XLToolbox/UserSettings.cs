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

        new public static UserSettings Default
        {
            get
            {
                return _lazy.Value;
            }
        }

        private static Lazy<UserSettings> _lazy = new Lazy<UserSettings>(() =>
        {
            UserSettings s = FromFileOrDefault<UserSettings>(UserSettingsFile);
            Bovender.UserSettings.UserSettingsBase.Default = s;
            return s;
        });

        #endregion

        #region User settings

        /// <summary>
        /// Wraps the singleton PresetsRepository's Presets property.
        /// </summary>
        public ObservableCollection<Preset> ExportPresets
        {
            get
            {
                return PresetsRepository.Default.Presets;
            }
            set
            {
                PresetsRepository.Default.Presets = value;
            }
        }

        /// <summary>
        /// Gets or sets the last used export preset. This property is exempt from serialization.
        /// On getting the ExportPreset, the PresetsRepository is searched for the last used
        /// preset's checksum hash. If no corresponding preset is found, it returns null.
        /// </summary>
        [YamlDotNet.Serialization.YamlIgnore]
        public Preset ExportPreset
        {
            get
            {
                _exportPreset = PresetsRepository.Default.FindByHash(_exportPresetHash);
                return _exportPreset;
            }
            set
            {
                _exportPreset = value;
            }
        }

        /// <summary>
        /// Gets or sets the checksum hash of the last used export preset.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Upon serialization
        /// to YAML, the hash will be computed on the fly from the actual last used export
        /// preset object. Upon deserialization, the hash will be stored in a private string
        /// field, and is used to retrieve the corresponding Preset object from the
        /// PresetsRepository collection when the ExportPreset property is accessed.
        /// </para>
        /// <para>
        /// This special handling serves to prevent YamlDotNet from writing object references,
        /// which are not really human-friendly. The Preset objects in the user settings file
        /// are all in the ExportPresets collection, and the last used preset is only stored
        /// using its checksum hash.
        /// </para>
        /// </remarks>
        public string ExportPresetHash
        {
            get
            {
                if (_exportPreset != null)
                {
                    _exportPresetHash = _exportPreset.ComputeMD5Hash();
                }
                else
                {
                    _exportPresetHash = null;
                }
                return _exportPresetHash;
            }
            set
            {
                _exportPresetHash = value;
            }
        }

        /// <summary>
        /// Gets or sets the checksum hash of the batch export settings' Preset.
        /// Upon serialization (get), this hash is computed from the current
        /// BatchExportSettings. Upon deserialization (set), the hash is stored
        /// and used when the BatchExportSettings is accessed.
        /// </summary>
        public string BatchExportPresetHash
        {
            get
            {
                if (_batchExportSettings != null && _batchExportSettings.Preset != null)
                {
                    _batchExportPresetHash = BatchExportSettings.Preset.ComputeMD5Hash();
                }
                else
                {
                    _batchExportPresetHash = null;
                }
                return _batchExportPresetHash;
            }
            set
            {
                _batchExportPresetHash = value;
            }
        }

        public BatchExportSettings BatchExportSettings
        {
            get
            {
                if (_batchExportSettings != null)
                {
                    _batchExportSettings.Preset = PresetsRepository.Default.FindByHash(_batchExportPresetHash);
                    // Invalidate the entire batch export settings if the
                    // export preset is invalid.
                    if (_batchExportSettings.Preset == null)
                    {
                        _batchExportSettings = null;
                    }
                }
                return _batchExportSettings;
            }
            set
            {
                _batchExportSettings = value;
            }
        }

        public Unit ExportUnit
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

        public bool WorksheetManagerAlwaysOnTop { get; set; }

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

        public int UpdateCheckInterval
        {
            get
            {
                if (_updateCheckInterval <= 0)
                {
                    _updateCheckInterval = 7;
                }
                return _updateCheckInterval;
            }
            set
            {
                _updateCheckInterval = value;
            }
        }

        public DateTime LastUpdateCheck
        {
            get
            {
                if (_lastUpdateCheck == null)
                {
                    _lastUpdateCheck = new DateTime(2016, 1, 1);
                }
                return _lastUpdateCheck;
            }
            set {
                _lastUpdateCheck = value;
            }
        }

        public string LastVersionSeen
        {
            get
            {
                if (_lastVersionSeen == null)
                {
                    _lastVersionSeen = "0.0.0";
                }
                return _lastVersionSeen;
            }
            set
            {
                _lastVersionSeen = value;
            }
        }

        #endregion

        #region Overrides

        public override string GetSettingsFilePath()
        {
            return UserSettingsFile;
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

        private string _exportPath;
        private Preset _exportPreset;
        private string _exportPresetHash;
        private BatchExportSettings _batchExportSettings;
        private string _batchExportPresetHash;
        private Unit _lastExportUnit;
        private Csv.CsvFile _csvImport;
        private Csv.CsvFile _csvExport;
        private int _taskPaneWidth;
        private DateTime _lastUpdateCheck;
        private int _updateCheckInterval;
        private string _lastVersionSeen;

        #endregion
    }
}

/* BatchExportSettings.cs
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
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using XLToolbox.WorkbookStorage;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// Holds settings for a specific batch export process.
    /// </summary>
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    public class BatchExportSettings : Settings
    {
        #region Factory

        /// <summary>
        /// Returns the BatchExportSettings object that was last used
        /// and stored in the assembly's Properties, or Null if no
        /// stored object exists.
        /// </summary>
        /// <returns>Last stored BatchExportSettings object, or Null
        /// if no such object exists.</returns>
        public static BatchExportSettings FromLastUsed()
        {
            return UserSettings.UserSettings.Default.BatchExportSettings;
        }

        /// <summary>
        /// Returns the BatchExportSettings object that was last used
        /// and stored in the workbookContext's hidden storage area,
        /// or the one stored in the assembly's Properties, or Null if no
        /// stored object exists.
        /// </summary>
        /// <param name="workbookContext"></param>
        /// <returns>Last stored BatchExportSettings object, or Null
        /// if no such object exists.</returns>
        public static BatchExportSettings FromLastUsed(Workbook workbookContext)
        {
            using (Store store = new Store(workbookContext))
            {
                BatchExportSettings settings = store.Get<BatchExportSettings>(
                    typeof(BatchExportSettings).ToString()
                    );
                if (settings != null && settings.Preset != null)
                {
                    // Replace the Preset object in the settings with the equivalent
                    // one from the PresetsRepository, or add
                    // it to the repository if no Preset with the same checksum hash exists.
                    settings.Preset = PresetsRepository.Default.FindOrAdd(settings.Preset);
                    return settings;
                }
                else
                {
                    return BatchExportSettings.FromLastUsed();
                }
            }
        }

        #endregion

        #region Public properties

        public BatchExportScope Scope { get; set; }
        public BatchExportObjects Objects { get; set; }
        public BatchExportLayout Layout { get; set; }
        public String Path { get; set; }

        #endregion

        #region Constructors

        public BatchExportSettings()
            : base()
        {
            Preset = PresetsRepository.Default.First;
        }

        public BatchExportSettings(Preset preset)
            : this()
        {
            Preset = preset;
        }

        #endregion

        #region Public methods

        public void Store()
        {
            UserSettings.UserSettings.Default.BatchExportSettings = this;
        }

        public void Store(Workbook workbookContext)
        {
            Store();
            using (Store store = new Store(workbookContext))
            {
                store.Put<BatchExportSettings>(
                    typeof(BatchExportSettings).ToString(),
                    this);
            }
        }

        #endregion
    }
}

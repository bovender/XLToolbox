/* PresetsRepository.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
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
using System.IO.IsolatedStorage;
using System.Xml.Serialization;
using System.Collections.ObjectModel;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// Repository for export settings, is concerned with storing and
    /// retrieving a collection of <see cref="Presets"/>.
    /// </summary>
    [Serializable]
    public class PresetsRepository : IDisposable
    {
        #region Public properties

        public ObservableCollection<Preset> Presets
        {
            get
            {
                ObservableCollection<Preset> presets = UserSettings.Default.ExportPresets as ObservableCollection<Preset>;
                if (presets == null)
                {
                    presets = new ObservableCollection<Preset>();
                    UserSettings.Default.ExportPresets = presets;
                }
                return presets;
            }
            set
            {
                UserSettings.Default.ExportPresets = value;
            }
        }

        #endregion

        #region Constructor

        public PresetsRepository()
            : base ()
        { }

        /// <summary>
        /// Creates a new Presets repository, loads previously saved presets
        /// and adds the <paramref name="addPreset"/> to the repository.
        /// </summary>
        /// <param name="addPreset">Preset to add to the repository.</param>
        public PresetsRepository(Preset addPreset)
            : this()
        {
            Presets.Add(addPreset);
        }

        #endregion

        #region Add and remove

        public void Add(Preset exportSettings)
        {
            Presets.Add(exportSettings);
        }

        public void Remove(Preset exportSettings)
        {
            Presets.Remove(exportSettings);
        }

        #endregion

        #region Disposal

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }

        ~PresetsRepository()
        {
            Dispose(false);
        }

        #endregion

        #region Private fields

        bool _disposed;

        #endregion
    }
}

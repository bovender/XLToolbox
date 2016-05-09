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
    public class PresetsRepository
    {
        #region Singleton factory

        /// <summary>
        /// The default (singleton) presets repository will always have at least
        /// one Preset, unless the Presets collection is replaced with an empty
        /// collection.
        /// </summary>
        public static PresetsRepository Default
        {
            get
            {
                return _lazy.Value;
            }
            set
            {
                _lazy = new Lazy<PresetsRepository>(() => value);
            }
        }

        private static Lazy<PresetsRepository> _lazy = new Lazy<PresetsRepository>(() => new PresetsRepository());

        #endregion

        #region Public properties

        /// <summary>
        /// Gets or sets the Presets collection. 
        /// </summary>
        public ObservableCollection<Preset> Presets { get; set; }

        #endregion

        #region Public methods

        public void Add(Preset preset)
        {
            Presets.Add(preset);
        }

        public void Remove(Preset preset)
        {
            Presets.Remove(preset);
        }

        /// <summary>
        /// Retrieves a Preset by its MD5 hash. If the hash is not
        /// found in the repository, this method returns null.
        /// </summary>
        /// <param name="hash">MD5 hash to look up.</param>
        /// <returns>Corresponding Preset, or null if no Preset with
        /// this hash exists.</returns>
        public Preset FindByHash(string hash)
        {
            if (String.IsNullOrWhiteSpace(hash))
            {
                return null;
            }
            else
            {
                return Presets.FirstOrDefault(p => p.ComputeMD5Hash() == hash);
            }
        }

        /// <summary>
        /// Looks up a preset by its guid and returns the Preset object
        /// stored in the repository. If the guid is not found, the Preset
        /// is added to the repository.
        /// </summary>
        /// <remarks>
        /// This method serves to reuse existing Presets: Given a Preset
        /// object, if a Preset object with the same guid exists in the
        /// repository, this object will be returned to that the original
        /// Preset object can be discarded.
        /// </remarks>
        public Preset FindOrAdd(Preset preset)
        {
            if (preset == null)
            {
                throw new ArgumentNullException();
            }
            Preset existingPreset = FindByHash(preset.ComputeMD5Hash());
            if (existingPreset == null)
            {
                Add(preset);
                return preset;
            }
            else
            {
                return existingPreset;
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor, calls the base constructor, does nothing more.
        /// </summary>
        private PresetsRepository()
            : base()
        {
            Presets = new ObservableCollection<Preset>() { new Preset() };
        }

        #endregion
    }
}

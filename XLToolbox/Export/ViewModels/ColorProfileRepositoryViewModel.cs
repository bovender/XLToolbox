/* ColorProfileRepositoryViewModel.cs
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
using Bovender.Mvvm.ViewModels;
using XLToolbox.Export.Models;
using System.Collections.ObjectModel;

namespace XLToolbox.Export.ViewModels
{
    using ColorProfileViewModelCollection = ObservableCollection<ColorProfileViewModel>;

    /// <summary>
    /// View model for a color profile repository, i.e. the collection
    /// of color profiles installed on the system.
    /// </summary>
    public class ColorProfileRepositoryViewModel : ViewModelBase
    {
        #region Public properties

        /// <summary>
        /// Gets or sets the current color space. Setting the color space
        /// switches the exposed color profile collection.
        /// </summary>
        public ColorSpace ColorSpace
        {
            get
            {
                return _colorSpace;
            }
            set
            {
                _colorSpace = value;
                OnPropertyChanged("ColorSpace");
                OnPropertyChanged("Profiles");
            }
        }

        public ObservableCollection<ColorProfileViewModel> Profiles
        {
            get
            {
                ObservableCollection<ColorProfileViewModel> profiles;
                if (_profiles.TryGetValue(ColorSpace, out profiles))
                {
                    return profiles;
                }
                else
                {
                    profiles = new ObservableCollection<ColorProfileViewModel>();
                    _profiles.Add(ColorSpace, profiles);
                    return profiles;
                }
            }
        }

        public ColorProfileViewModel SelectedProfile
        {
            get
            {
                ColorProfileViewModel vm;
                if (_selectedProfiles.TryGetValue(ColorSpace, out vm))
                {
                    return vm;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                _selectedProfiles[ColorSpace] = value;
                OnPropertyChanged("SelectedProfile");
            }
        }

        #endregion

        #region Public methods

        public bool HasProfilesForColorSpace(ColorSpace colorSpace)
        {
            ColorProfileViewModelCollection c;
            if (_profiles.TryGetValue(colorSpace, out c))
            {
                return c.Count > 0;
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region Constructor

        public ColorProfileRepositoryViewModel()
            : base()
        {
            _profiles = new Dictionary<ColorSpace, ColorProfileViewModelCollection>();
            _selectedProfiles = new Dictionary<ColorSpace, ColorProfileViewModel>();
            BuildCollections();
        }

        #endregion

        #region Implementation of ViewModelBase

        public override object RevealModelObject()
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Private methods

        private void BuildCollections()
        {
            string dir = Bovender.Unmanaged.Pinvoke.GetColorDirectory();
            if (String.IsNullOrEmpty(dir))
            {
                throw new InvalidOperationException(
                    "Windows did not tell color profile directory");
            }
            ColorProfileViewModel vm;
            ColorProfileViewModelCollection coll;
            foreach (string fn in System.IO.Directory.EnumerateFiles(
                System.IO.Path.Combine(dir, "*.ics")))
            {
                vm = ColorProfileViewModel.CreateFromFile(fn);
                if (vm != null)
                {
                    if (!_profiles.TryGetValue(vm.ColorSpace, out coll))
                    {
                        coll = new ColorProfileViewModelCollection();
                        _profiles.Add(vm.ColorSpace, coll);
                    }
                    coll.Add(vm);
                }
            }
            OnPropertyChanged("Profiles");
        }

        #endregion

        #region Private fields

        /// <summary>
        /// A dictionary of color profile collection, one per color space.
        /// </summary>
        private Dictionary<ColorSpace, ColorProfileViewModelCollection> _profiles;

        /// <summary>
        /// A dictionary of selected profiles, one per color space.
        /// </summary>
        private Dictionary<ColorSpace, ColorProfileViewModel> _selectedProfiles;

        /// <summary>
        /// Holds the currently selected color space.
        /// </summary>
        private ColorSpace _colorSpace;

        #endregion
    }
}

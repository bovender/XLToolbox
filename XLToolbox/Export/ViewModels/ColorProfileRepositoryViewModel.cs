/* ColorProfileRepositoryViewModel.cs
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
                _updating = true;
                _colorSpace = value;
                // OnPropertyChanged("ColorSpace");
                OnPropertyChanged("Profiles");
                _updating = false;
                OnPropertyChanged("SelectedProfile");
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

        /// <summary>
        /// Gets or sets the selected profile for the current color space.
        /// If no profile was previously selected, gets the first profile
        /// for the color space that exists.
        /// </summary>
        public ColorProfileViewModel SelectedProfile
        {
            get
            {
                // To make any potentially subscribed WPF ComboBox behave
                // correctly, we *must* provide a null value as SelectedItem
                // when the ItemsSource collection changes.
                if (_updating) return null;

                ColorProfileViewModel vm;
                if (_selectedProfiles.TryGetValue(ColorSpace, out vm))
                {
                    return vm;
                }
                else
                {
                    if (Profiles.Count > 0)
                    {
                        _selectedProfiles[ColorSpace] = Profiles[0];
                        return _selectedProfiles[ColorSpace];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            set
            {
                // Value may be null if bound view is a WPF ComboBox that
                // will set the SelectedItem property to null when the
                // ItemsSource property is updated.
                if (value != null)
                {
                    _selectedProfiles[ColorSpace] = value;
                }
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

        /// <summary>
        /// Selects a color profile for the current color space by name,
        /// if it exists.
        /// </summary>
        /// <param name="colorProfile">Color profile to select in the
        /// current color space.</param>
        /// <returns>True if the profile is exists and was selected,
        /// false if not.</returns>
        public bool SelectIfExists(string colorProfile)
        {
            ColorProfileViewModel vm = Profiles.FirstOrDefault(
                c => String.Equals(c.Name, colorProfile));
            if (vm != null)
            {
                SelectedProfile = vm;
                return true;
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

            using (Bovender.Unmanaged.DllManager dllManager = new Bovender.Unmanaged.DllManager())
            {
                ColorProfileViewModel vm;
                ColorProfileViewModelCollection coll;
                dllManager.LoadDll("lcms2.dll");
                foreach (string fn in System.IO.Directory.EnumerateFiles(dir,
                    "*" + Lcms.Constants.COLOR_PROFILE_EXTENSION))
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

        private bool _updating;

        #endregion
    }
}

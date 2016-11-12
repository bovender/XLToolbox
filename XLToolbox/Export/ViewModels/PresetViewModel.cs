/* PresetViewModel.cs
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
using Microsoft.Office.Interop.Excel;
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using XLToolbox.Export.Models;
using XLToolbox.WorkbookStorage;
using System.ComponentModel;
using System.Xml.Serialization;
using System.IO;
using Bovender.Extensions;

namespace XLToolbox.Export.ViewModels
{
    /// <summary>
    /// View model for graphic export settings.
    /// </summary>
    public class PresetViewModel : ViewModelBase
    {
        #region Factory

        public static PresetViewModel FromLastUsed()
        {
            Preset p = Preset.FromLastUsed();
            if (p != null)
            {
                return new PresetViewModel(p);
            }
            else
            {
                return null;
            }
        }
        
        public static PresetViewModel FromLastUsed(Workbook workbookContext)
        {
            Preset p = Preset.FromLastUsed(workbookContext);
            if (p != null)
            {
                return new PresetViewModel(p);
            }
            else
            {
                return null;
            }
        }

        #endregion

        #region Properties

        public string Name
        {
            get { return _preset.Name; }
            set
            {
                _preset.Name = value;
                _customName = true;
                OnPropertyChanged("Name");
                OnPropertyChanged("DisplayString");
            }
        }

        public int Dpi
        {
            get { return _preset.Dpi; }
            set
            {
                _preset.Dpi = value;
                UpdateName();
                OnPropertyChanged("Dpi");
            }
        }

        public int DpiMinimum { get { return 100;  } }
        public int DpiMaximum { get { return 1200;  } }
        public int DpiIncrements { get { return 50; } }

        public bool IsDpiEnabled
        {
            get
            {
                return !_preset.IsVectorType;
            }
        }

        public EnumProvider<FileType> FileType
        {
            get
            {
                if (_fileTypeProvider == null)
                {
                    _fileTypeProvider = new EnumProvider<FileType>();
                    _fileTypeProvider.AsEnum = _preset.FileType;
                    _fileTypeProvider.PropertyChanged +=
                        (object sender, PropertyChangedEventArgs args) =>
                        {
                            _preset.FileType = _fileTypeProvider.AsEnum;
                            UpdateCmykEnabledState();
                            SanitizeSettings();
                            OnPropertyChanged("FileType." + args.PropertyName);
                            OnPropertyChanged("IsColorSpaceEnabled");
                            OnPropertyChanged("IsDpiEnabled");
                            OnPropertyChanged("IsTransparencyEnabled");
                            UpdateName();
                        };
                }
                return _fileTypeProvider;
            }
        }

        public ColorSpaceProvider ColorSpace
        {
            get
            {
                if (_colorSpaceProvider == null)
                {
                    _colorSpaceProvider = new ColorSpaceProvider();
                    _colorSpaceProvider.AsEnum = _preset.ColorSpace;
                    _colorSpaceProvider.PropertyChanged +=
                        (object sender, PropertyChangedEventArgs args) =>
                        {
                            if (args.PropertyName == "AsEnum")
                            {
                                _preset.ColorSpace = _colorSpaceProvider.AsEnum;

                                // Filter the color profiles according to color space.
                                ColorProfiles.ColorSpace = _colorSpaceProvider.AsEnum;

                                if (ColorProfiles.Profiles.Count == 0)
                                {
                                    _preset.UseColorProfile = false;
                                }

                                // If CMYK is selected, color management *must* be used.
                                // Note that CMYK is disabled automatically if there is not
                                // CMYK color profile installed; however, due to a bug in the
                                // WPF combobox, CMYK could still be selected using type-ahead
                                // search.
                                if (_colorSpaceProvider.AsEnum == Models.ColorSpace.Cmyk)
                                {
                                    _preset.UseColorProfile = true;
                                    this.Transparency.AsEnum = Models.Transparency.WhiteCanvas;
                                    _mustUseColorProfile = true;
                                }
                                else
                                {
                                    _mustUseColorProfile = false;
                                }
                                SanitizeSettings();

                                OnPropertyChanged("IsUseColorProfileEnabled");
                                OnPropertyChanged("UseColorProfile");
                                OnPropertyChanged("IsColorProfilesEnabled");
                                OnPropertyChanged("IsTransparencyEnabled");
                                UpdateName();
                            }
                        };
                }
                return _colorSpaceProvider;
            }
        }

        public bool IsColorSpaceEnabled
        {
            get
            {
                return !_preset.IsVectorType;
            }
        }

        public TransparencyProvider Transparency
        {
            get
            {
                if (_transparencyProvider == null)
                {
                    _transparencyProvider = new TransparencyProvider();
                    _transparencyProvider.AsEnum = _preset.Transparency;
                    _transparencyProvider.PropertyChanged +=
                        (object sender, PropertyChangedEventArgs args) =>
                        {
                            _preset.Transparency = _transparencyProvider.AsEnum;
                            OnPropertyChanged("Transparency." + args.PropertyName);
                            UpdateName();
                        };
                }
                return _transparencyProvider;
            }
        }

        public bool IsTransparencyEnabled
        {
            get
            {
                return !_preset.IsVectorType && ColorSpace.AsEnum.SupportsTransparency();
            }
        }

        public string ToolTip
        {
            get
            {
                return _preset.GetDefaultName();
            }
        }

        public bool UseColorProfile
        {
            get
            {
                return _preset.UseColorProfile;
            }
            set
            {
                _preset.UseColorProfile = value;
                UpdateName();
                OnPropertyChanged("UseColorProfile");
                OnPropertyChanged("IsColorProfilesEnabled");
            }
        }

        public bool IsUseColorProfileEnabled
        {
            get
            {
                return (ColorProfiles.Profiles.Count > 0)
                    && !_preset.IsVectorType
                    && !_mustUseColorProfile;
            }
            protected set
            {
                _mustUseColorProfile = value;
                OnPropertyChanged("IsUseColorProfileEnabled");
            }
        }

        public ColorProfileRepositoryViewModel ColorProfiles
        {
            get
            {
                if (_colorProfiles == null)
                {
                    _colorProfiles = new ColorProfileRepositoryViewModel();
                    _colorProfiles.PropertyChanged += (sender, args) =>
                    {
                        if (args.PropertyName == "SelectedProfile")
                        {
                            ColorProfileViewModel cpvm = ColorProfiles.SelectedProfile;
                            _preset.ColorProfile = (cpvm != null) ? cpvm.Name : String.Empty;
                        }
                        OnPropertyChanged("ColorProfiles" + args.PropertyName);
                    };
                }
                return _colorProfiles;
            }
        }

        public bool IsColorProfilesEnabled
        {
            get
            {
                return UseColorProfile && (ColorProfiles.Profiles.Count > 0)
                    && !_preset.IsVectorType;
            }
        }

        #endregion

        #region Overrides

        public override string DisplayString
        {
            get
            {
                return Name;
            }
            set
            {
                Name = value;
            }
        }

        #endregion

        #region Constructors

        public PresetViewModel(Preset preset)
            : base()
        {
            _preset = preset;
            if (_preset != null)
            {
                _customName = !String.Equals(Name, _preset.GetDefaultName());
            }
            ColorProfiles.ColorSpace = _preset.ColorSpace;
            UpdateCmykEnabledState();
            if (!ColorProfiles.SelectIfExists(_preset.ColorProfile))
            {
                UseColorProfile = false;
            }
        }

        public PresetViewModel()
            : this(new Preset())
        { }

        #endregion

        #region Public methods

        public void Store(Workbook workbookContext)
        {
            _preset.Store(workbookContext);
        }

        public void Store()
        {
            _preset.Store();
        }

        #endregion

        #region Private methods

        private void UpdateName()
        {
            if (!_customName)
            {
                Name = _preset.GetDefaultName();
                _customName = false;
            }
            OnPropertyChanged("ToolTip");
        }

        /// <summary>
        /// Enables or disables the CMYK color space depending on the selected
        /// file type.
        /// </summary>
        private void UpdateCmykEnabledState()
        {
            ColorSpace.GetViewModel(Models.ColorSpace.Cmyk).IsEnabled =
                ColorProfiles.HasProfilesForColorSpace(Models.ColorSpace.Cmyk) &&
                _preset.FileType.SupportsCmyk();
        }

        /// <summary>
        /// Ensures that only valid options are chosen by falling back to defaults
        /// if necessary
        /// </summary>
        private void SanitizeSettings()
        {
            if (ColorSpace.AsEnum == Models.ColorSpace.Cmyk && !_preset.FileType.SupportsCmyk())
            {
                // Fall back to a default colorspace
                ColorSpace.AsEnum = Models.ColorSpace.Rgb;
            }
            // Transparency is a yes/no decision, no need to change the setting per se
            // The combobox is disabled if the file type/color space do not support
            // transparency, and the actual value is not important it is not supported
            // if (!(Transparency.AsEnum == Models.Transparency.WhiteCanvas) && !IsTransparencyEnabled)
            // {
            //     Transparency.AsEnum = Models.Transparency.WhiteCanvas;
            // }
        }

        #endregion

        #region Private fields

        Preset _preset;
        ColorSpaceProvider _colorSpaceProvider;
        EnumProvider<FileType> _fileTypeProvider;
        TransparencyProvider _transparencyProvider;
        bool _customName;
        bool _mustUseColorProfile;
        ColorProfileRepositoryViewModel _colorProfiles;

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _preset;
        }

        #endregion
    }
}

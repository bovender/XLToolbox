using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;
using XLToolbox.Export.Models;
using System.ComponentModel;

namespace XLToolbox.Export.ViewModels
{
    /// <summary>
    /// View model for graphic export settings.
    /// </summary>
    public class PresetViewModel : ViewModelBase
    {
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
                            _preset.ColorSpace = _colorSpaceProvider.AsEnum;
                            OnPropertyChanged("ColorSpace." + args.PropertyName);
                            UpdateName();
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
                return !_preset.IsVectorType;
            }
        }

        public string Tooltip
        {
            get
            {
                return _preset.GetDefaultName();
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

        public PresetViewModel()
            : base()
        {
            _preset = new Preset();
        }

        public PresetViewModel(Preset preset)
            : base()
        {
            _preset = preset;
            _customName = !String.Equals(Name, _preset.GetDefaultName());
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
            OnPropertyChanged("Tooltip");
        }

        #endregion

        #region Private fields

        Preset _preset;
        ColorSpaceProvider _colorSpaceProvider;
        EnumProvider<FileType> _fileTypeProvider;
        TransparencyProvider _transparencyProvider;
        bool _customName;

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _preset;
        }

        #endregion
    }
}

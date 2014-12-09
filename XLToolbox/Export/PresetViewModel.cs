using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Export
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

        public bool IsDpiEnabled
        {
            get
            {
                return !_preset.IsVectorType;
            }
        }

        public FileType FileType
        {
            get { return _preset.FileType; }
            set
            {
                _preset.FileType = value;
                UpdateName();
                OnPropertyChanged("FileType");
                OnPropertyChanged("IsColorSpaceEnabled");
                OnPropertyChanged("IsDpiEnabled");
            }
        }

        public ColorSpace ColorSpace
        {
            get { return _preset.ColorSpace; }
            set
            {
                _preset.ColorSpace = value;
                UpdateName();
                OnPropertyChanged("ColorSpace");
            }
        }

        public bool IsColorSpaceEnabled
        {
            get
            {
                return !_preset.IsVectorType;
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
        }

        #endregion

        #region Private fields

        Preset _preset;
        bool _customName;

        #endregion

        #region Implemenation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _preset;
        }

        #endregion
    }
}

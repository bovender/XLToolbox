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
    public class SettingsViewModel : ViewModelBase
    {
        #region Properties

        public string Name
        {
            get { return _exportSettings.Name; }
            set
            {
                _exportSettings.Name = value;
                _customName = true;
                OnPropertyChanged("Name");
            }
        }

        public int Dpi
        {
            get { return _exportSettings.Dpi; }
            set
            {
                _exportSettings.Dpi = value;
                UpdateName();
                OnPropertyChanged("Dpi");
            }
        }

        public bool IsDpiEnabled
        {
            get
            {
                return !_exportSettings.IsVectorType;
            }
        }

        public FileType FileType
        {
            get { return _exportSettings.FileType; }
            set
            {
                _exportSettings.FileType = value;
                UpdateName();
                OnPropertyChanged("FileType");
                OnPropertyChanged("IsColorSpaceEnabled");
                OnPropertyChanged("IsDpiEnabled");
            }
        }

        public ColorSpace ColorSpace
        {
            get { return _exportSettings.ColorSpace; }
            set
            {
                _exportSettings.ColorSpace = value;
                UpdateName();
                OnPropertyChanged("ColorSpace");
            }
        }

        public bool IsColorSpaceEnabled
        {
            get
            {
                return !_exportSettings.IsVectorType;
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

        #region Constructor

        public SettingsViewModel(Settings exportSettings)
            : base()
        {
            _exportSettings = exportSettings;
            _customName = !String.Equals(Name, _exportSettings.GetDefaultName());
        }

        #endregion

        #region Private methods

        private void UpdateName()
        {
            if (!_customName)
            {
                Name = _exportSettings.GetDefaultName();
                _customName = false;
            }
        }

        #endregion

        #region Private fields

        Settings _exportSettings;
        bool _customName;

        #endregion

        #region Implemenation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _exportSettings;
        }

        #endregion
    }
}

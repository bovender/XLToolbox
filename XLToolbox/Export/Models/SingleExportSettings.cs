using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// Holds settings for a specific single export process.
    /// </summary>
    public class SingleExportSettings : Settings
    {
        #region Public properties

        public double Width
        {
            get { return _width; }
            set
            {
                // Preserve aspect, but only if current dimensions are not 0
                if (PreserveAspect && _width != 0 && _height != 0)
                {

                    double _aspect = _height / _width;
                    _height = value * _aspect;
                }
                _width = value;
            }
        }

        public double Height
        {
            get { return _height; }
            set
            {
                if (PreserveAspect && _width != 0 && _height != 0)
                {
                    double _aspect = _width / _height;
                    _width = value * _aspect;
                }
                _height = value;
            }
        }

        public bool PreserveAspect { get; set; }

        #endregion

        #region Constructors

        public SingleExportSettings()
            : base()
        { }

        public SingleExportSettings(Preset preset, double width, double height, bool preserveAspect)
            : this()
        {
            Preset = preset;
            Width = width;
            Height = height;
            PreserveAspect = preserveAspect;
        }

        #endregion

        #region Implementation of Settings

        public override void Store()
        {
            Properties.Settings.Default.LastSingleExportSetting = this;
            Properties.Settings.Default.Save();
        }

        #endregion

        #region Private fields

        double _width;
        double _height;

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// Model for graphic export settings.
    /// </summary>
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    public class Preset 
    {
        #region Properties

        public string Name { get; set; }
        public int Dpi { get; set; }
        public FileType FileType { get; set; }
        public ColorSpace ColorSpace { get; set; }
        public bool IsVectorType
        {
            get
            {
                return FileType == FileType.Emf || FileType == FileType.Svg;
            }
        }
        public int Bpp
        {
            get
            {
                return ColorSpace.ToBPP();
            }
        }

        public Transparency Transparency { get; set; }

        #endregion

        #region Constructors

        public Preset()
        {
            Name = GetDefaultName();
        }

        public Preset(FileType fileType, int dpi, ColorSpace colorSpace)
        {
            FileType = fileType;
            Dpi = dpi;
            ColorSpace = colorSpace;
            GetDefaultName();
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Returns a default name for the current settings that
        /// is created from the individual properties.
        /// </summary>
        /// <returns></returns>
        public string GetDefaultName()
        {
            if (IsVectorType)
            {
                return FileType.ToString();
            }
            else
            {
                ViewModels.ColorSpaceProvider csp = new ViewModels.ColorSpaceProvider();
                ViewModels.TransparencyProvider tp = new ViewModels.TransparencyProvider();
                csp.AsEnum = ColorSpace;
                tp.AsEnum = Transparency;
                return String.Format("{0}, {1} dpi, {2}, {3}",
                    FileType.ToString(), Dpi, csp.AsString, tp.AsString);
            }
        }

        #endregion

        #region Overrides

        public override string ToString()
        {
            return GetDefaultName();
        }

        #endregion
    }
}

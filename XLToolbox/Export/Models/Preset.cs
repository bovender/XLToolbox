using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// Model for graphic export settings.
    /// </summary>
    [Serializable]
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
                return String.Format("{0}, {1} dpi, {2}",
                    FileType.ToString(), Dpi, ColorSpace.ToString());
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

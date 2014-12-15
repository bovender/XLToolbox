using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FreeImageAPI;

namespace XLToolbox.Export
{
    public enum FileType
    {
        Tiff,
        Png,
        Svg,
        Emf
    }

    static class FileTypeExtensions
    {
        public static FREE_IMAGE_FORMAT ToFreeImageFormat(this FileType fileType)
        {
            switch (fileType)
            {
                case FileType.Tiff: return FREE_IMAGE_FORMAT.FIF_TIFF;
                case FileType.Png: return FREE_IMAGE_FORMAT.FIF_PNG;
                default:
                    throw new NotImplementedException("No FREE_IMAGE_FORMAT match for " + fileType.ToString());
            }
        }
    }
}

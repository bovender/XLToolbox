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

    public static class FileTypeExtensions
    {
        public static FREE_IMAGE_FORMAT ToFreeImageFormat(this FileType fileType)
        {
            switch (fileType)
            {
                case FileType.Tiff: return FREE_IMAGE_FORMAT.FIF_TIFF;
                case FileType.Png: return FREE_IMAGE_FORMAT.FIF_PNG;
                default:
                    throw new InvalidOperationException("No FREE_IMAGE_FORMAT match for " + fileType.ToString());
            }
        }

        public static string ToFileNameExtension(this FileType fileType)
        {
            switch (fileType)
            {
                case FileType.Tiff: return ".tif";
                case FileType.Png: return ".png";
                default:
                    throw new InvalidOperationException("No file name extension defined for " + fileType.ToString());
            }
        }

        public static bool SupportsTransparency(this FileType fileType)
        {
            switch (fileType)
            {
                case FileType.Tiff: return true;
                case FileType.Png: return true;
                default:
                    throw new InvalidOperationException("Transparency support unknown for " + fileType.ToString());
            }
        }
    }
}

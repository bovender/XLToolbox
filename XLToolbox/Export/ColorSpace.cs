using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FreeImageAPI;

namespace XLToolbox.Export
{
    // TODO: Add CMYK colorspace.
    public enum ColorSpace
    {
        Monochrome,
        GrayScale,
        Rgb,
        // Cmyk
    }

    public static class ColorSpaceExtensions
    {
        public static FREE_IMAGE_COLOR_TYPE ToFreeImageColorType(this ColorSpace colorSpace)
        {
            switch (colorSpace)
            {
                case ColorSpace.Rgb: return FREE_IMAGE_COLOR_TYPE.FIC_RGBALPHA;
                case ColorSpace.GrayScale: return FREE_IMAGE_COLOR_TYPE.FIC_RGBALPHA;
                case ColorSpace.Monochrome: return FREE_IMAGE_COLOR_TYPE.FIC_PALETTE;
                default:
                    throw new InvalidOperationException(
                        "No FREE_IMAGE_COLOR_TYPE match for " + colorSpace.ToString());
            }
        }

        public static FREE_IMAGE_COLOR_DEPTH ToFreeImageColorDepth(this ColorSpace colorSpace)
        {
            switch (colorSpace)
            {
                case ColorSpace.Monochrome: return FREE_IMAGE_COLOR_DEPTH.FICD_01_BPP;
                case ColorSpace.Rgb: return FREE_IMAGE_COLOR_DEPTH.FICD_32_BPP;
                // case ColorSpace.Cmyk: return FREE_IMAGE_COLOR_DEPTH.FICD_32_BPP;
                case ColorSpace.GrayScale:
                    return FREE_IMAGE_COLOR_DEPTH.FICD_FORCE_GREYSCALE | FREE_IMAGE_COLOR_DEPTH.FICD_08_BPP;
                default:
                    throw new InvalidOperationException(
                        "No FREE_IMAGE_COLOR_DEPTH match for " + colorSpace.ToString());
            }
        }

        public static int ToBPP(this ColorSpace colorSpace)
        {
            switch (colorSpace)
            {
                case ColorSpace.Monochrome: return 1;
                case ColorSpace.Rgb: return 24;
                // case ColorSpace.Cmyk: return 32;
                case ColorSpace.GrayScale: return 8;
                default:
                    throw new InvalidOperationException(
                        "BPP not defined for " + colorSpace.ToString());
            }
        }
    }
}

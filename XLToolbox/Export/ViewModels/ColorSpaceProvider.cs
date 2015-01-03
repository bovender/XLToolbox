using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm;
using XLToolbox.Export.Models;

namespace XLToolbox.Export.ViewModels
{
    public class ColorSpaceProvider : EnumProvider<ColorSpace>
    {
        protected override string GetDescription(ColorSpace member)
        {
            switch (member)
            {
                case ColorSpace.GrayScale: return Strings.GrayScale;
                case ColorSpace.Monochrome: return Strings.Monochrome;
                case ColorSpace.Rgb: return Strings.Rgb;
                default:
                    throw new InvalidOperationException(
                        "No localized description available for " + member.ToString());
            }
        }
    }
}

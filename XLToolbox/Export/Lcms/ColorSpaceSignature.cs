/* ColorSpaceSignature.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLToolbox.Export.Models;

namespace XLToolbox.Export.Lcms
{
    /// <summary>
    /// Select color spaces from [LCMS2-2.6]/include/lcms2.h.
    /// </summary>
    public enum ColorSpaceSignature
    {
        SigRgbData = 0x52474220,
        SigGrayData = 0x47524159,
        SigCmykData = 0x434D594B,
    }

    public static class cmsColorSpaceSignatureExtension
    {
        /// <summary>
        /// Converts an LCMS ColorSpaceSignature to an XL Toolbox ColorSpace.
        /// </summary>
        /// <param name="signature">ColorSpaceSignature to convert.</param>
        /// <returns>Corresponding ColorSpace.</returns>
        /// <exception cref="InvalidOperationException">if there is no
        /// conversion defined for the given ColorSpaceSignature.</exception>
        public static ColorSpace ToColorSpace(this ColorSpaceSignature signature)
        {
            switch (signature)
            {
                case ColorSpaceSignature.SigCmykData:
                    return ColorSpace.Cmyk;
                case ColorSpaceSignature.SigGrayData:
                    return ColorSpace.GrayScale;
                case ColorSpaceSignature.SigRgbData:
                    return ColorSpace.Rgb;
                default:
                    throw new InvalidOperationException(
                        "No conversion defined for " + signature.ToString());
            }
        }

        /// <summary>
        /// Converts an XL Toolbox ColorSpace to an LCMS ColorSpaceSignature.
        /// </summary>
        /// <param name="colorSpace"></param>
        /// <param name="colorSpace">Corresponding ColorSpaceSignature.</param>
        /// <returns></returns>
        public static ColorSpaceSignature FromColorSpace(this ColorSpaceSignature signature,
            ColorSpace colorSpace)
        {
            switch (colorSpace)
            {
                case ColorSpace.Cmyk:
                    return ColorSpaceSignature.SigCmykData;
                case ColorSpace.GrayScale:
                    return ColorSpaceSignature.SigGrayData;
                case ColorSpace.Rgb:
                    return ColorSpaceSignature.SigRgbData;
                default:
                    throw new InvalidOperationException(
                        "No conversion defined for " + colorSpace.ToString());
            }
        }
    }
}

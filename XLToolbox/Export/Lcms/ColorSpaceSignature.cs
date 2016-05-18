/* ColorSpaceSignature.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2016 Daniel Kraus
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
        XYZ = 0x58595A20,
        Lab = 0x4C616220,
        Luv = 0x4C757620,
        YCbCr = 0x59436272,
        Yxy = 0x59787920,
        Rgb = 0x52474220,
        Gray = 0x47524159,
        Hsv = 0x48535620,
        Hls = 0x484C5320,
        Cmyk = 0x434D594B,
        Cmy = 0x434D5920,
        MCH1 = 0x4D434831,
        MCH2 = 0x4D434832,
        MCH3 = 0x4D434833,
        MCH4 = 0x4D434834,
        MCH5 = 0x4D434835,
        MCH6 = 0x4D434836,
        MCH7 = 0x4D434837,
        MCH8 = 0x4D434838,
        MCH9 = 0x4D434839,
        MCHA = 0x4D43483A,
        MCHB = 0x4D43483B,
        MCHC = 0x4D43483C,
        MCHD = 0x4D43483D,
        MCHE = 0x4D43483E,
        MCHF = 0x4D43483F,
        Named = 0x6e6d636c,
        // 1color = 0x31434C52,
        // 2color = 0x32434C52,
        // 3color = 0x33434C52,
        // 4color = 0x34434C52,
        // 5color = 0x35434C52,
        // 6color = 0x36434C52,
        // 7color = 0x37434C52,
        // 8color = 0x38434C52,
        // 9color = 0x39434C52,
        // 10color = 0x41434C52,
        // 11color = 0x42434C52,
        // 12color = 0x43434C52,
        // 13color = 0x44434C52,
        // 14color = 0x45434C52,
        // 15color = 0x46434C52,
        LuvK = 0x4C75764B,
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
                case ColorSpaceSignature.Cmyk:
                    return ColorSpace.Cmyk;
                case ColorSpaceSignature.Gray:
                    return ColorSpace.GrayScale;
                case ColorSpaceSignature.Rgb:
                    return ColorSpace.Rgb;
                default:
                    throw new NotImplementedException(
                        "LCMS signature to color space conversion for " +
                        signature.ToString().ToUpper() +
                        " not implemented.");
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
                    return ColorSpaceSignature.Cmyk;
                case ColorSpace.GrayScale:
                    return ColorSpaceSignature.Gray;
                case ColorSpace.Rgb:
                    return ColorSpaceSignature.Rgb;
                default:
                    throw new NotImplementedException(
                        "Color space to LCMS signature conversion for " +
                        colorSpace.ToString().ToUpper() +
                        " not implemented.");
            }
        }
    }
}

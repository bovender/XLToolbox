/* Formatters.cs
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

namespace XLToolbox.Export.Lcms
{
    /// <summary>
    /// Defines constants in analogy to the compiler macros found
    /// in lcms2.h.
    /// </summary>
    internal class Formatters
    {
        #region Public formatter constants

        public const UInt32 GRAY_8 = // (COLORSPACE_SH(PT_GRAY)|CHANNELS_SH(1)|BYTES_SH(1))
            PT_GRAY << COLORSPACE | 1 << CHANNELS | 1 << BYTES;

        public const UInt32 TYPE_RGB_8 = // (COLORSPACE_SH(PT_RGB)|CHANNELS_SH(3)|BYTES_SH(1))
            PT_RGB << COLORSPACE | 3 << CHANNELS | 1 << BYTES;

        public const UInt32 TYPE_BGR_8 = // (COLORSPACE_SH(PT_RGB)|CHANNELS_SH(3)|BYTES_SH(1)|DOSWAP_SH(1))
            PT_RGB << COLORSPACE | 3 << CHANNELS | 1 << BYTES | 1 << DOSWAP;

        public const UInt32 TYPE_RGBA_8 = // (COLORSPACE_SH(PT_RGB)|EXTRA_SH(1)|CHANNELS_SH(3)|BYTES_SH(1))
            PT_RGB << COLORSPACE | 1 << EXTRA | 3 << CHANNELS | 1 << BYTES;

        public const UInt32 TYPE_BGRA_8 = // (COLORSPACE_SH(PT_RGB)|EXTRA_SH(1)|CHANNELS_SH(3)|BYTES_SH(1)|DOSWAP_SH(1)|SWAPFIRST_SH(1))
            PT_RGB << COLORSPACE | 1 << EXTRA | 3 << CHANNELS | 1 << BYTES | 1 << DOSWAP | 1 << SWAPFIRST;

        public const UInt32 TYPE_CMYK = // (COLORSPACE_SH(PT_CMYK)|CHANNELS_SH(4)|BYTES_SH(1))
            PT_CMYK << COLORSPACE | 4 << CHANNELS | 1 << BYTES;

        #endregion

        #region Bitshift constants

        private const int COLORSPACE = 16;
        private const int SWAPFIRST = 14;
        private const int DOSWAP = 10;
        private const int EXTRA = 7;
        private const int CHANNELS = 3;
        private const int BYTES = 0;

        #endregion

        #region Color constants
        private const int PT_RGB = 4;
        private const int PT_GRAY = 3;
        private const int PT_CMYK = 6;

        #endregion
    }
}

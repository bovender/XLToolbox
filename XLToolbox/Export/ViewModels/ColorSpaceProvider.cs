/* ColorSpaceProvider.cs
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

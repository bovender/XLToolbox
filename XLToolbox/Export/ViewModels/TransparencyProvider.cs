/* TransparencyProvider.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using Bovender.Mvvm;
using XLToolbox.Export.Models;

namespace XLToolbox.Export.ViewModels
{
    public class TransparencyProvider : EnumProvider<Transparency>
    {
        protected override string GetDescription(Transparency member)
        {
            switch (member)
            {
                case Transparency.TransparentCanvas: return Strings.TransparentCanvas;
                case Transparency.TransparentWhite: return Strings.TransparentWhite;
                case Transparency.WhiteCanvas: return Strings.WhiteCanvas;
                default:
                    throw new InvalidOperationException(
                        "No localized description available for " + member.ToString());
            }
        }

        protected override string GetTooltip(Transparency member)
        {
            switch (member)
            {
                case Transparency.TransparentCanvas: return Strings.TransparentCanvasTooltip;
                case Transparency.TransparentWhite: return Strings.TransparentWhiteTooltip;
                case Transparency.WhiteCanvas: return Strings.WhiteCanvasTooltip;
                default:
                    throw new InvalidOperationException(
                        "No localized tooltip available for " + member.ToString());
            }
        }
    }
}

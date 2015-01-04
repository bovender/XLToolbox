/* FileType.cs
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
using FreeImageAPI;
using System.Windows.Data;
using System.ComponentModel;

namespace XLToolbox.Export.Models
{
    public enum FileType
    {
        [Description("TIFF")]
        Tiff,
        [Description("PNG")]
        Png,
        // [Description("SVG")]
        // Svg,
        [Description("EMF")]
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
                case FileType.Emf: return ".emf";
                default:
                    throw new InvalidOperationException("No file name extension defined for " + fileType.ToString());
            }
        }

        public static string ToFileFilter(this FileType fileType)
        {
            string result;
            switch (fileType)
            {
                case FileType.Emf:  result = Strings.EmfFiles; break;
                case FileType.Png:  result = Strings.PngFiles; break;
                // case FileType.Svg:  result = Strings.SvgFiles; break;
                case FileType.Tiff: result = Strings.TifFiles; break;
                default:
                    throw new InvalidOperationException(
                        "No file filter defined for " + fileType.ToString());
            }
            return result + "|*" + fileType.ToFileNameExtension();
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

/* PathExtensions.cs
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
using System.IO;

namespace Bovender
{
    public static class PathHelpers
    {
        /// <summary>
        /// Extracts the directory information from <paramref name="path"/>.
        /// If <paramref name="path"/> is an existing directory, the entire
        /// <paramref name="path"/> will be returned. Otherwise, the return
        /// value of <see cref="System.IO.Path.GetDirectoryName()"/> will
        /// be returned.
        /// </summary>
        /// <param name="path">Path to examine.</param>
        /// <returns><paramref name="path"/> or result of
        /// <see cref="System.IO.Path.GetFileName()"/>.</returns>
        public static string GetDirectoryPart(string path)
        {
            if (System.IO.Directory.Exists(path))
            {
                return path;
            }
            else
            {
                return System.IO.Path.GetDirectoryName(path);
            }
        }

        /// <summary>
        /// Extracts the file name from <paramref name="path"/>.
        /// If <paramref name="path"/> is an existing directory, an
        /// empty string will be returned. Otherwise, the return
        /// value of <see cref="System.IO.Path.GetFileName()"/> will
        /// be returned.
        /// </summary>
        /// <param name="path">Path to examine.</param>
        /// <returns>String.Empty or return value of
        /// <see cref="System.IO.Path.GetFileName()"/>.</returns>
        public static string GetFileNamePart(string path)
        {
            if (System.IO.Directory.Exists(path))
            {
                return String.Empty;
            }
            else
            {
                return System.IO.Path.GetFileName(path);
            }
        }
    }
}

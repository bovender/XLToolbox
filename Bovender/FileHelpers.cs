/* FileHelpers.cs
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
using System.Security.Cryptography;

namespace Bovender
{
    public static class FileHelpers
    {
        /// <summary>
        /// Computes the Sha1 hash of a given file.
        /// </summary>
        /// <param name="file">File to compute the Sha1 for.</param>
        /// <returns>Sha1 hash.</returns>
        public static string Sha1Hash(string file)
        {
            using (FileStream fs = new FileStream(file, FileMode.Open))
            using (BufferedStream bs = new BufferedStream(fs))
            {
                using (SHA1Managed sha1 = new SHA1Managed())
                {
                    byte[] hash = sha1.ComputeHash(bs);
                    StringBuilder formatted = new StringBuilder(2 * hash.Length);
                    foreach (byte b in hash)
                    {
                        formatted.AppendFormat("{0:x2}", b);
                    }
                    return formatted.ToString();
                }
            }
        }

        /// <summary>
        /// Returns the name of the directory contained in path. Unlike
        /// System.IO.Path.GetDirectoryName(), this function does not simply
        /// strip the part after the last path separator from the path, but
        /// rather performs a check if the path is in fact an existing
        /// directory (without file name added). If the path does not exist
        /// (neither file nor directory), System.IO.Path.GetDirectoryName()
        /// is called to produce the result.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string GetDirectoryName(string path)
        {
            if (Directory.Exists(path))
            {
                return path;
            }
            else {
                return Path.GetDirectoryName(path);
            }
        }
    }
}

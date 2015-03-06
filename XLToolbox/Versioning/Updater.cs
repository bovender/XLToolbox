/* Updater.cs
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
using System.Text.RegularExpressions;

namespace XLToolbox.Versioning
{
    public class Updater : Bovender.Versioning.Updater
    {
        #region Overrides

        protected override Uri GetVersionInfoUri()
        {
            return new Uri(Properties.Settings.Default.VersionInfoUrl);
        }

        protected override Bovender.Versioning.SemanticVersion GetCurrentVersion()
        {
            return XLToolbox.Versioning.SemanticVersion.CurrentVersion();
        }

        protected override string BuildDestinationFileName()
        {
            string fn;
            Regex r = new Regex(@"(?<fn>[^/]+?exe)");
            Match m = r.Match(DownloadUri.ToString());
            if (m.Success)
            {
                fn = m.Groups["fn"].Value;
            }
            else
            {
                fn = String.Format("XL_Toolbox_{0}.exe", NewVersion.ToString());
            };
            return Path.Combine(DestinationFolder, fn);
        }

        #endregion
    }
}

/* SemanticVersion.cs
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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace XLToolbox.Versioning
{
    public class SemanticVersion : Bovender.Versioning.SemanticVersion
    {
        #region Current version singleton

        public static SemanticVersion Current
        {
            get
            {
                return _lazy.Value;
            }
        }

        private static readonly Lazy<SemanticVersion> _lazy =
            new Lazy<SemanticVersion>(() => new SemanticVersion());

        #endregion

        #region Public method

        public string BrandName
        {
            get
            {
                return Properties.Settings.Default.AddinName + " " + ToString();
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Creates an instance with the current XL Toolbox version
        /// </summary>
        private SemanticVersion() : base(Assembly.GetExecutingAssembly()) { }

        public SemanticVersion(string version) : base(version) { }

        #endregion
    }
}

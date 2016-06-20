/* SemanticVersion.cs
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
using System.Reflection;

namespace XLToolbox.Versioning
{
    public class SemanticVersion : Bovender.Versioning.SemanticVersion
    {
        #region Static 'overrides'

        /// <summary>
        /// Returns the current version of the XL Toolbox addin.
        /// </summary>
        /// <returns></returns>
        new public static Bovender.Versioning.SemanticVersion CurrentVersion()
        {
            if (_currentVersion == null)
            {
                _currentVersion = Bovender.Versioning.SemanticVersion.CurrentVersion(
                    Assembly.GetExecutingAssembly()
                    );
                Logger.Info("Current version: {0}", _currentVersion);
            }
            return _currentVersion;
        }

        #endregion

        #region Static methods

        public static string BrandName
        {
            get
            {
                return Properties.Settings.Default.AddinName + " " + CurrentVersion().ToString();
            }
        }

        #endregion

        #region Private static fields

        private static Bovender.Versioning.SemanticVersion _currentVersion;

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

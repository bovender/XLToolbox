/* Updater.cs
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
using Bovender.Versioning;
using Bovender.Mvvm.Actions;
using Bovender.Extensions;

namespace XLToolbox.Versioning
{
    public class Updater : Bovender.Versioning.Updater
    {
        #region Static properties
        
        public static Updater Default { get; set; }

        public static bool CanCheck { get; set; } // TODO: Find a way around this global variable
        
        #endregion

        #region Static event
        
        public static event EventHandler<EventArgs> CanCheckChanged;

        #endregion

        #region Public static methods

        public static Updater CreateDefault(IReleaseInfo releaseInfo)
        {
            Default = new Updater(releaseInfo);
            return Default;
        }
        
        #endregion

        #region Private static methods

        private static void OnCanCheckChanged()
        {
            EventHandler<EventArgs> h = CanCheckChanged;
            if (h != null)
            {
                h(null, new EventArgs());
            }
        }
        
        #endregion

        #region Constructor

        public Updater(Bovender.Versioning.IReleaseInfo releaseInfo)
            : base(releaseInfo)
        {
            CurrentVersion = SemanticVersion.Current;
        }
        
        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

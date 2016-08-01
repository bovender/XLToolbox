/* ReleaseInfoViewModel.cs
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

namespace XLToolbox.Versioning
{
    public class ReleaseInfoViewModel : Bovender.Versioning.ReleaseInfoViewModel
    {
        #region Constructors

        public ReleaseInfoViewModel()
            : base(new ReleaseInfo(), SemanticVersion.Current)
        {
            UpdateAvailableMessage.Sent += (sender, args) =>
            {
                Logger.Info("UpdateAvailableMessage received");
                UpdaterViewModel updaterVM = new UpdaterViewModel(Updater.CreateDefault(ReleaseInfo));
                updaterVM.ShowUpdateAvailableView();
            };
            NoUpdateAvailableMessage.Sent += (sender, args) =>
            {
                Logger.Info("NoUpdateAvailableMessage received");
                Updater.CanCheck = false;
            };
            ExceptionMessage.Sent += (sender, args) =>
            {
                Logger.Warn("Exception during update check");
                Logger.Warn(Exception);
                Updater.CanCheck = true;
            };
        }
        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

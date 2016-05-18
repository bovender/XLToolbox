/* UpdaterViewModel.cs
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
using Bov = Bovender.Versioning;

namespace XLToolbox.Versioning
{
    /// <summary>
    /// Provides a singleton instance of Bovender.UpdaterViewModel.
    /// </summary>
    public class UpdaterViewModel
    {
        #region Singleton factory

        private static readonly Lazy<Bovender.Versioning.UpdaterViewModel> _instance =
            new Lazy<Bovender.Versioning.UpdaterViewModel>(
                () => new Bovender.Versioning.UpdaterViewModel(new Updater())
            );

        public static Bovender.Versioning.UpdaterViewModel Instance
        {
            get
            {
                return _instance.Value;
            }
        }

        #endregion
    }
}

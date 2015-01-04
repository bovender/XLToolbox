/* WindowSettings.cs
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
using System.Configuration;
using System.Windows;

namespace Bovender.Mvvm.Views.Settings
{
    public class WindowSettings : ApplicationSettingsBase
    {
        #region Constructor

        public WindowSettings(Window window) : base(window.ToString()) { }

        #endregion

        #region Properties

        /// <summary>
        /// The rectangle describing the window's coordinates.
        /// </summary>
        [UserScopedSetting]
        public Rect Rect
        {
            get
            {
                // Cannot return this["Rect"] as Rect
                // because Rect is not nullable
                object obj = this["Rect"];
                if (obj == null) {
                    return Rect.Empty;
                }
                else {
                    return (Rect)obj;
                }
            }
            set
            {
                this["Rect"] = value;
            }
        }

        [UserScopedSetting]
        public System.Windows.WindowState State
        {
            get
            {
                object obj = this["State"];
                if (obj == null)
                {
                    return System.Windows.WindowState.Normal;
                }
                else
                {
                    return (System.Windows.WindowState)obj;
                }
            }
            set
            {
                this["State"] = value;
            }
        }

        [UserScopedSetting]
        public System.Windows.Forms.Screen Screen
        {
            get
            {
                return (System.Windows.Forms.Screen)this["Screen"];
            }
            set
            {
                this["Screen"] = value;
            }
        }

        #endregion

    }
}

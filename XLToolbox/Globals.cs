using Microsoft.Office.Tools;
/* Globals.cs
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

[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("Tests,PublicKey=00240000048000009400000006020000002400005253413100040000010001000fd238794ea4b062ff5ce8b89502543e868a4c1b032b9f6d99f1b77646d5c79e197f5d1f3786c6f8712d7654315382ac8e71775daa68d4124eb2afad3b8b784115095a64a44921cf7c3ed437d6f0b1a99949ce7a26ceda5d2e84f18c7597b428d566712cfb18f8ca4ebff3b537b22acda01f7486c6821fd535197cb443747dc0")]

namespace XLToolbox
{
    public static class Globals
    {
        public static CustomTaskPaneCollection CustomTaskPanes { get; set; }

        /// <summary>
        /// XAML-accessible name of the add-in.
        /// </summary>
        public static string AddinName
        {
            get
            {
                return Properties.Settings.Default.AddinName;
            }
        }

        /// <summary>
        /// XAML-accessible website URL.
        /// </summary>
        public static Uri WebsiteUri
        {
            get
            {
                return new Uri(Properties.Settings.Default.WebsiteUrl);
            }
        }
    }
}

/* CsvFile.cs
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
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using XLToolbox.Excel;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace XLToolbox.Csv
{
    /// <summary>
    /// Provides import/export settings and methods for CSV files.
    /// </summary>
    public abstract class CsvFileBase : Bovender.Mvvm.Models.ProcessModel
    {
        #region Properties

        public string FileName { get; set; }

        public string FieldSeparator
        {
            get
            {
                return Settings.FieldSeparator;
            }
            set
            {
                Settings.FieldSeparator = value;
            }
        }

        public string DecimalSeparator
        {
            get
            {
                return Settings.DecimalSeparator;
            }
            set
            {
                Settings.DecimalSeparator = value;
            }
        }

        public string ThousandsSeparator
        {
            get
            {
                return Settings.ThousandsSeparator;
            }
            set
            {
                Settings.ThousandsSeparator = value;
            }
        }

        public bool Tabularize
        {
            get
            {
                return Settings.Tabularize;
            }
            set
            {
                Settings.Tabularize = value;
            }
        }

        public CsvSettings Settings
        {
            get
            {
                if (_settings == null)
                {
                    _settings = new CsvSettings();
                }
                return _settings;
            }
            set
            {
                _settings = value;
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Instantiates the class with a new CsvSettings object.
        /// </summary>
        protected CsvFileBase()
            : base()
        { }

        /// <summary>
        /// Instantiates the class with a given CsvSettings object.
        /// </summary>
        /// <param name="settings">Settings to initialize the
        /// class with.</param>
        protected CsvFileBase(CsvSettings settings)
            : this()
        {
            Settings = settings;
        }

        #endregion

        #region Protected property

        protected NumberFormatInfo NumberFormatInfo
        {
            get
            {
                return Settings.NumberFormatInfo;
            }
        }

        #endregion

        #region Private fields

        private CsvSettings _settings;

        #endregion
    }
}

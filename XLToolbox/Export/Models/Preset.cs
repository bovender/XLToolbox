/* Preset.cs
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
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using XLToolbox.Excel.Instance;
using XLToolbox.WorkbookStorage;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// Model for graphic export settings.
    /// </summary>
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    public class Preset 
    {
        #region Factory

        public static Preset FromLastUsed()
        {
            return Properties.Settings.Default.ExportPreset;
        }

        public static Preset FromLastUsed(Workbook workbookContext)
        {
            Store store = new Store(workbookContext);
            Preset preset = store.Get<Preset>(typeof(Preset).ToString());
            if (preset != null)
            {
                return preset;
            }
            else
            {
                return Preset.FromLastUsed();
            }
        }

        #endregion

        #region Properties

        public string Name { get; set; }
        public int Dpi { get; set; }
        public FileType FileType { get; set; }
        public ColorSpace ColorSpace { get; set; }
        public bool IsVectorType
        {
            get
            {
                return FileType == FileType.Emf; // || FileType == FileType.Svg;
            }
        }
        public int Bpp
        {
            get
            {
                return ColorSpace.ToBPP();
            }
        }
        public Transparency Transparency { get; set; }
        public bool UseColorProfile { get; set; }
        public string ColorProfile { get; set; }

        #endregion

        #region Constructors

        public Preset()
        {
            FileType = Models.FileType.Png;
            ColorSpace = Models.ColorSpace.Rgb;
            Transparency = Models.Transparency.TransparentCanvas;
            Dpi = 300;
            Name = GetDefaultName();
        }

        public Preset(FileType fileType, int dpi, ColorSpace colorSpace)
        {
            FileType = fileType;
            Dpi = dpi;
            ColorSpace = colorSpace;
            GetDefaultName();
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Returns a default name for the current settings that
        /// is created from the individual properties.
        /// </summary>
        /// <returns></returns>
        public string GetDefaultName()
        {
            if (IsVectorType)
            {
                return FileType.ToString();
            }
            else
            {
                ViewModels.ColorSpaceProvider csp = new ViewModels.ColorSpaceProvider();
                ViewModels.TransparencyProvider tp = new ViewModels.TransparencyProvider();
                csp.AsEnum = ColorSpace;
                tp.AsEnum = Transparency;
                return String.Format("{0}, {1} dpi, {2}, {3}",
                    FileType.ToString(), Dpi, csp.SelectedItem.DisplayString, tp.SelectedItem.DisplayString);
            }
        }
        
        public void Store(Workbook workbookContext)
        {
            Properties.Settings.Default.ExportPreset = this;
            Properties.Settings.Default.Save();
            using (Store store = new Store(workbookContext))
            {
                store.Put<Preset>(typeof(Preset).ToString(), this);
            }
        }

        public void Store()
        {
            Store(ExcelInstance.Application.ActiveWorkbook);
        }

        #endregion

        #region Overrides

        public override string ToString()
        {
            return GetDefaultName();
        }

        #endregion
    }
}

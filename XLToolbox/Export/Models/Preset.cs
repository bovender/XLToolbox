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
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Text;
using XLToolbox.Excel.ViewModels;
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
            return UserSettings.Default.RetrieveExportPresetOrDefault();
        }

        public static Preset FromLastUsed(Workbook workbookContext)
        {
            ObservableCollection<Preset> repo = UserSettings.Default.ExportPresets;
            Store store = new Store(workbookContext);
            if (repo != null && repo.Count > 0)
            {
                int index = store.Get(Properties.StoreNames.Default.ExportPreset,
                    UserSettings.Default.LastExportPreset, 0, repo.Count - 1);
                return repo[index];
            }
            else
            {
                if (repo == null)
                {
                    repo = new ObservableCollection<Preset>();
                    UserSettings.Default.ExportPresets = repo;
                }
                Preset p = new Preset();
                repo.Add(p);
                return p;
            }
        }

        #endregion

        #region Properties

        public string Name { get; set; }
        public int Dpi { get; set; }
        public FileType FileType { get; set; }
        public ColorSpace ColorSpace { get; set; }

        [YamlDotNet.Serialization.YamlIgnore]
        public bool IsVectorType
        {
            get
            {
                return FileType == FileType.Emf; // || FileType == FileType.Svg;
            }
        }
        
        [YamlDotNet.Serialization.YamlIgnore]
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
                // Construct some EnumProviders to get nice text representations.
                ViewModels.ColorSpaceProvider csp = new ViewModels.ColorSpaceProvider();
                ViewModels.TransparencyProvider tp = new ViewModels.TransparencyProvider();
                csp.AsEnum = ColorSpace;
                tp.AsEnum = Transparency;

                string cp = String.Empty;
                if (UseColorProfile)
                {
                    cp = ", " + Strings.ColorManagement;
                }

                return String.Format("{0}, {1} dpi, {2}, {3}{4}",
                    FileType.ToString(),
                    Dpi,
                    csp.SelectedItem.DisplayString,
                    tp.SelectedItem.DisplayString,
                    cp);
            }
        }
        
        public void Store(Workbook workbookContext)
        {
            UserSettings.Default.StoreExportPreset(this);
            using (Store store = new Store(workbookContext))
            {
                store.Put(Properties.StoreNames.Default.ExportPreset,
                    UserSettings.Default.GetExportPresetIndex(this));
            }
        }

        public void Store()
        {
            Store(Excel.ViewModels.Instance.Default.ActiveWorkbook);
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

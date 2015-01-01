using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using XLToolbox.WorkbookStorage;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// Holds settings for a specific batch export process.
    /// </summary>
    [Serializable]
    [SettingsSerializeAs(SettingsSerializeAs.Xml)]
    public class BatchExportSettings : Settings
    {
        #region Factory

        /// <summary>
        /// Returns the BatchExportSettings object that was last used
        /// and stored in the assembly's Properties, or Null if no
        /// stored object exists.
        /// </summary>
        /// <returns>Last stored BatchExportSettings object, or Null
        /// if no such object exists.</returns>
        public static BatchExportSettings FromLastUsed()
        {
            return Properties.Settings.Default.BatchExportSettings;
        }

        /// <summary>
        /// Returns the BatchExportSettings object that was last used
        /// and stored in the workbookContext's hidden storage area,
        /// or the one stored in the assembly's Properties, or Null if no
        /// stored object exists.
        /// </summary>
        /// <param name="workbookContext"></param>
        /// <returns>Last stored BatchExportSettings object, or Null
        /// if no such object exists.</returns>
        public static BatchExportSettings FromLastUsed(Workbook workbookContext)
        {
            Store store = new Store(workbookContext);
            BatchExportSettings settings = store.Get<BatchExportSettings>(
                typeof(BatchExportSettings).ToString()
                );
            if (settings != null)
            {
                return settings;
            }
            else
            {
                return BatchExportSettings.FromLastUsed();
            }
        }

        #endregion

        #region Public properties

        public BatchExportScope Scope { get; set; }
        public BatchExportObjects Objects { get; set; }
        public BatchExportLayout Layout { get; set; }
        public String Path { get; set; }

        #endregion

        #region Constructors

        public BatchExportSettings()
            : base()
        { }

        public BatchExportSettings(Preset preset)
            : this()
        {
            Preset = preset;
        }

        #endregion

        #region Public methods

        public void Store()
        {
            Properties.Settings.Default.BatchExportSettings = this;
            Properties.Settings.Default.Save();
        }

        public void Store(Workbook workbookContext)
        {
            Store();
            Store store = new Store(workbookContext);
            store.Put<BatchExportSettings>(
                typeof(BatchExportSettings).ToString(),
                this);
        }

        #endregion
    }
}

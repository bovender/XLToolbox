using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Export
{
    /// <summary>
    /// Holds settings for a specific batch export process.
    /// </summary>
    public class BatchExportSettings : Settings
    {
        #region Factory

        public static BatchExportSettings FromStored()
        {
            return Properties.Settings.Default.LastBatchExportSetting;
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

        #region Implementation of Settings

        public override void Store()
        {
            Properties.Settings.Default.LastBatchExportSetting = this;
            Properties.Settings.Default.Save();
        }

        #endregion
    }
}

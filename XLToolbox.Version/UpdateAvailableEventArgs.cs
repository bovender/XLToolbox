using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Version
{
    public class UpdateAvailableEventArgs : EventArgs
    {
        public SemanticVersion NewVersion { get; set; }
        public string NewVersionInfo { get; set; }
        public Uri DownloadUrl { get; set; }
        public UpdateAvailableEventArgs(SemanticVersion newVersion, string newVersionInfo,
            Uri downloadUrl)
        {
            NewVersion = newVersion;
            NewVersionInfo = newVersionInfo;
            DownloadUrl = downloadUrl;
        }
    }
}

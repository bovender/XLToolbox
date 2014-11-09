using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Versioning
{
    public class Updater : Bovender.Versioning.Updater
    {
        protected override Uri GetVersionInfoUri()
        {
            return new Uri(Properties.Settings.Default.VersionInfoUrl);
        }

        protected override Bovender.Versioning.SemanticVersion CurrentVersion()
        {
            return XLToolbox.Versioning.SemanticVersion.CurrentVersion();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

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

        protected override string BuildDestinationFileName()
        {
            string fn;
            Regex r = new Regex(@"(?<fn>[^/]+?exe)");
            Match m = r.Match(DownloadUri.ToString());
            if (m.Success)
            {
                fn = m.Groups["fn"].Value;
            }
            else
            {
                fn = String.Format("XL_Toolbox_{0}.exe", NewVersion.ToString());
            };
            return Path.Combine(DestinationFolder, fn);
        }
    }
}

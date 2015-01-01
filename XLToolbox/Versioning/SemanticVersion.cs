using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace XLToolbox.Versioning
{
    public class SemanticVersion : Bovender.Versioning.SemanticVersion
    {
        #region Static 'overrides'

        /// <summary>
        /// Returns the current version of the XL Toolbox addin.
        /// </summary>
        /// <returns></returns>
        new public static Bovender.Versioning.SemanticVersion CurrentVersion()
        {
            return Bovender.Versioning.SemanticVersion.CurrentVersion(
                Assembly.GetExecutingAssembly()
                );
        }

        #endregion

        public static string BrandName
        {
            get
            {
                return Properties.Settings.Default.AddinName + " " + CurrentVersion().ToString();
            }
        }
    }
}

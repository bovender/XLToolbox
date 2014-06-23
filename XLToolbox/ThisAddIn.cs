using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using XLToolbox.Version;
using System.Threading;

namespace XLToolbox
{
    public partial class ThisAddIn
    {
        public Updater Updater { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
#if !DEBUG
            Globals.Ribbons.Ribbon.GroupDebug.Visible = false;
#endif
            GreetUser();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (Updater != null)
            {
                Updater.InstallUpdate();
            }
        }

        private void GreetUser()
        {
            SemanticVersion lastSeenVersion = new SemanticVersion(
                Properties.Settings.Default.LastVersionSeen);
            SemanticVersion currentVersion = SemanticVersion.CurrentVersion();
            if (currentVersion > lastSeenVersion)
            {
                WindowGreeter w = new WindowGreeter();
                w.Show();
                Properties.Settings.Default.LastVersionSeen = currentVersion.ToString();
                Properties.Settings.Default.Save();
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

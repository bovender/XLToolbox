using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;

namespace XLToolbox.Version
{
    /// <summary>
    /// Fetches version information from the internet and raises an UpdateAvailable
    /// event if a new version is available for download.
    /// </summary>
    public class Updater
    {
        private const string VERSIONINFOURL = "http://xltoolbox.sourceforge.net/version-ng.txt";
        private UpdateAvailableEventArgs UpdateArgs { get; set; }
        public event EventHandler<UpdateAvailableEventArgs> UpdateAvailable;
        public event EventHandler<DownloadStringCompletedEventArgs> FetchingVersionFailed;

        /// <summary>
        /// Downloads the current version information file asynchronously from the project
        /// home page.
        /// </summary>
        public void FetchVersionInformation()
        {
            WebClient web = new WebClient();
            web.DownloadStringCompleted += web_DownloadStringCompleted;
            web.DownloadStringAsync(new Uri(VERSIONINFOURL));
        }

        /// <summary>
        /// Inspects the downloaded version information
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void web_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                StringReader r = new StringReader(e.Result);
                SemanticVersion v = new SemanticVersion(r.ReadLine());
                Uri url = new Uri(r.ReadLine());
                string info = r.ReadLine();

                // If a new version is available, raise the corresponding event.
                if (v != SemanticVersion.CurrentVersion())
                {
                    UpdateArgs = new UpdateAvailableEventArgs(v, info, url);
                    OnUpdateAvailable();
                }
            }
            else
            {
                // Raise an event that signals failure.
                OnFetchingVersionFailed(e);
            }
        }

        protected virtual void OnUpdateAvailable()
        {
            if (UpdateAvailable != null)
            {
                UpdateAvailable(this, UpdateArgs);
            }
        }

        protected virtual void OnFetchingVersionFailed(DownloadStringCompletedEventArgs e)
        {
            if (FetchingVersionFailed != null)
            {
                FetchingVersionFailed(this, e);
            }
        }
    }
}

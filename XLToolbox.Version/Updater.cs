using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Security.Cryptography;

namespace XLToolbox.Version
{
    /// <summary>
    /// Fetches version information from the internet and raises an UpdateAvailable
    /// event if a new version is available for download.
    /// </summary>
    /// <remarks>
    /// The current version information resides in a simple text file which contains
    /// four lines:              e.g.
    /// 1) Current version       7.0.0-alpha.1
    /// 2) Download URL          http://sourceforge.net/projects/xltoolbox/files/XL_Toolbox_7.0.0-alpha.1.exe
    /// 3) Sha1 of executable    1234abcd...
    /// 4) Version description   This is the first release of the next generation Toolbox
    /// </remarks>
    public class Updater
    {
        private const string VERSIONINFOURL = "http://xltoolbox.sourceforge.net/version-ng.txt";
        private UpdateAvailableEventArgs UpdateArgs { get; set; }
        private string FileName { get; set; }
        private string Sha1 { get; set; }

        /// <summary>
        /// Signals that an updated version is available for download.
        /// </summary>
        public event EventHandler<UpdateAvailableEventArgs> UpdateAvailable;

        /// <summary>
        /// Signals that the version information could not be downloaded from the internet.
        /// </summary>
        public event EventHandler<DownloadStringCompletedEventArgs> FetchingVersionFailed;

        /// <summary>
        /// Signals a change in the download process of the executable file. This event is
        /// chained from WebClient's event with the same name.
        /// </summary>
        public event EventHandler<DownloadProgressChangedEventArgs> DownloadProgressChanged;

        /// <summary>
        /// Signals that the new release has been downloaded, verified and is ready to install.
        /// </summary>
        public event EventHandler<UpdateAvailableEventArgs> DownloadInstallable;

        /// <summary>
        /// Signals that the downloaded file could not be verified.
        /// </summary>
        public event EventHandler<UpdateAvailableEventArgs> DownloadFailedVerification;

        /// <summary>
        /// Downloads the current version information file asynchronously from the project
        /// home page.
        /// </summary>
        public void FetchVersionInformation()
        {
            WebClient downloadTxt = new WebClient();
            downloadTxt.DownloadStringCompleted += downloadTxt_DownloadStringCompleted;
            downloadTxt.DownloadStringAsync(new Uri(VERSIONINFOURL));
        }

        /// <summary>
        /// Downloads the current release from the internet.
        /// </summary>
        public void DownloadUpdate(string fileName)
        {
            FileName = fileName;
            WebClient downloadExe = new WebClient();
            downloadExe.DownloadProgressChanged += downloadExe_DownloadProgressChanged;
            downloadExe.DownloadFileCompleted += downloadExe_DownloadFileCompleted;
            downloadExe.DownloadFileAsync(UpdateArgs.DownloadUrl, fileName);
        }

        public void InstallUpdate()
        {
            // Compute the SHA1 again so we know it's current.
            ComputeSha1();
            if (Sha1 == UpdateArgs.Sha1)
            {
                System.Diagnostics.Process.Start(FileName, "/UPDATE");
            }
            else
            {
                OnDownloadFailedVerification();
            }
        }

        void downloadExe_DownloadFileCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            ComputeSha1();
            if (Sha1 == UpdateArgs.Sha1)
            {
                OnDownloadInstallable();
            }
            else
            {
                OnDownloadFailedVerification();
                /* throw new DownloadCorruptException(String.Format(
                    "Checksum of downloaded file {0} does not match expected checksum {1}",
                    Sha1, UpdateArgs.Sha1)); */
            };
        }

        void downloadExe_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            if (DownloadProgressChanged != null)
            {
                DownloadProgressChanged(this, e);
            }
        }

        /// <summary>
        /// Inspects the downloaded version information.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void downloadTxt_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                StringReader r = new StringReader(e.Result);
                SemanticVersion v = new SemanticVersion(r.ReadLine());
                Uri url = new Uri(r.ReadLine());
                string sha1 = r.ReadLine();
                string info = r.ReadLine();

                // If a new version is available, raise the corresponding event.
                if (v != SemanticVersion.CurrentVersion())
                {
                    UpdateArgs = new UpdateAvailableEventArgs(v, info, url, sha1);
                    OnUpdateAvailable();
                }
            }
            else
            {
                // Raise an event that signals failure.
                OnFetchingVersionFailed(e);
            }
        }

        protected virtual void OnDownloadInstallable()
        {
            if (DownloadInstallable != null)
            {
                DownloadInstallable(this, UpdateArgs);
            }
        }

        protected virtual void OnDownloadFailedVerification()
        {
            if (DownloadFailedVerification != null)
            {
                DownloadFailedVerification(this, UpdateArgs);
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

        /// <summary>
        /// Computes the Sha1 hash of the downloaded file.
        /// </summary>
        /// <returns></returns>
        private void ComputeSha1()
        {
            using (FileStream fs = new FileStream(FileName, FileMode.Open))
            using (BufferedStream bs = new BufferedStream(fs))
            {
                using (SHA1Managed sha1 = new SHA1Managed())
                {
                    byte[] hash = sha1.ComputeHash(bs);
                    StringBuilder formatted = new StringBuilder(2 * hash.Length);
                    foreach (byte b in hash)
                    {
                        formatted.AppendFormat("{0:X2}", b);
                    }
                    Sha1 = formatted.ToString();
                }
            }
        }
    }
}

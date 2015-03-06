/* Updater.cs
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
using System;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Text;

namespace Bovender.Versioning
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
    public abstract class Updater
    {
        #region Public properties

        public string DestinationFolder { get; set; }

        public SemanticVersion NewVersion { get; protected set; }

        public SemanticVersion CurrentVersion { get { return GetCurrentVersion(); } }

        /// <summary>
        /// If true, an updated version is available for download.
        /// </summary>
        public bool IsUpdateAvailable { get; protected set; }

        /// <summary>
        /// Indicates whether an update has been downloaded and could be
        /// installed.
        /// </summary>
        public bool IsUpdatePending { get; protected set; }

        /// <summary>
        /// The URI of the remote file.
        /// </summary>
        public Uri DownloadUri { get; protected set; }

        /// <summary>
        /// Returns true if the Sha1 of the downloaded file matches
        /// the one in the version information file.
        /// </summary>
        public bool IsVerifiedDownload { get; protected set; }

        /// <summary>
        /// The Sha1 hash of the remote file as reported in the version
        /// info file.
        /// </summary>
        /// 
        public string UpdateSha1 { get; protected set; }

        /// <summary>
        /// Summary of changes as reported in the version info file.
        /// </summary>
        public string UpdateSummary { get; protected set; }

        /// <summary>
        /// Determines whether the current user is authorized to write to the folder
        /// where the addin files are stored. If the user does not have write permissions,
        /// he/she cannot update the addin by herself/hisself.
        /// </summary>
        public virtual bool IsAuthorized
        {
            get
            {
                string addinPath = AppDomain.CurrentDomain.BaseDirectory;
                /* Todo: compute permissions, rather than try and catch */
                try
                {
                    string fn = Path.Combine(addinPath, "xltbupd.test");
                    using (FileStream f = new FileStream(fn,
                        FileMode.Create, FileAccess.Write))
                    {
                        f.WriteByte(0xff);
                    };
                    File.Delete(fn);
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        public Exception DownloadException { get; protected set; }

        #endregion

        #region Events

        /// <summary>
        /// Signals that the current version information has been refreshed.
        /// </summary>
        public event EventHandler<EventArgs> CheckForUpdateFinished;

        /// <summary>
        /// Signals a change in the download process of the executable file. This event is
        /// chained from WebClient's event with the same name.
        /// </summary>
        public event EventHandler<DownloadProgressChangedEventArgs> DownloadProgressChanged;

        /// <summary>
        /// Signals that an update has been downloaded. Subscribers need to
        /// check if the update is actually installable.
        /// </summary>
        public event EventHandler<EventArgs> DownloadUpdateFinished;

        #endregion

        #region Public methods

        /// <summary>
        /// Downloads the current version information file asynchronously from the project
        /// home page.
        /// </summary>
        /// <remarks>
        /// Eventually triggers the UpdateAvailable or NoUpdateAvailable events if the current version
        /// information was downloaded successfully; and triggers the FetchingVersionFailed
        /// event if the version information could not be downloaded.
        /// </remarks>
        public void CheckForUpdate()
        {
            _versionInfoClient = new WebClient();
            _versionInfoClient.DownloadStringCompleted += VersionInfoClient_DownloadStringCompleted;
            _versionInfoClient.DownloadStringAsync(GetVersionInfoUri());
        }

        public void CancelCheckForUpdate()
        {
            if (_versionInfoClient != null)
            {
                _versionInfoClient.CancelAsync();
            }
        }

        /// <summary>
        /// Downloads the current release from the internet.
        /// </summary>
        public void DownloadUpdate()
        {
            _destinationFileName = BuildDestinationFileName();

            /* Check if the file exists already. If the Sha1 is identical,
             * do not download it again. If the Sha1 is different, it is a file
             * with the same name, but different content (broken download?).
             */
            if (File.Exists(_destinationFileName) &&
                FileHelpers.Sha1Hash(_destinationFileName) == UpdateSha1)
            {
                // Bypass the download and signal that the file is present
                IsVerifiedDownload = true;
                OnDownloadUpdateFinished();
            }
            else
            {
                _client = new WebClient();
                _client.DownloadProgressChanged += _client_DownloadProgressChanged;
                _client.DownloadFileCompleted += _client_DownloadFileCompleted;
                _client.DownloadFileAsync(DownloadUri, _destinationFileName);
            }
        }

        public void CancelDownload()
        {
            _client.CancelAsync();
        }

        /// <summary>
        /// Verifies the Sha1 checksum of the file on disk again and executes
        /// the file if it is valid. Does nothing if no update is available.
        /// </summary>
        /// <exception cref="DownloadCorruptException">if the Sha1 is unexpected</exception>
        public void InstallUpdate()
        {
            if (!IsUpdateAvailable) return;

            // As a security measure, compute the SHA1 again so we know it's current.
            IsVerifiedDownload = FileHelpers.Sha1Hash(_destinationFileName) == UpdateSha1;

            if (IsVerifiedDownload)
            {
                DoInstallUpdate();
            }
            else
            {
                throw new DownloadCorruptException("The Sha1 checksum of the file on disk is unexpected.");
            }
        }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Returns the URI for the file that provides current version information.
        /// </summary>
        /// <returns>URI for version info file.</returns>
        protected abstract Uri GetVersionInfoUri();

        /// <summary>
        /// Returns the version number of the current program.
        /// </summary>
        /// <returns>Instance of <see cref="SemanticVersion"/> representing
        /// the current version number.</returns>
        protected abstract SemanticVersion GetCurrentVersion();

        #endregion

        #region Protected virtual methods

        protected virtual void DoDownload()
        {
            string defaultPath = Properties.Versioning.Default.DownloadPath;
            if (string.IsNullOrEmpty(defaultPath))
            {
                defaultPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            };

        }

        /// <summary>
        /// Performs the actual download once the user has confirmed the download
        /// destination.
        /// </summary>
        protected virtual void ConfirmDownload()
        {

        }

        protected virtual bool CanDownload()
        {
            return IsAuthorized;
        }

        protected virtual void OnDownloadUpdateFinished()
        {
            if (DownloadUpdateFinished != null)
            {
                DownloadUpdateFinished(this, EventArgs.Empty);
            }
        }

        protected virtual void OnCheckForUpdateFinished()
        {
            if (CheckForUpdateFinished != null)
            {
                CheckForUpdateFinished(this, EventArgs.Empty);
            }
        }

        /// <summary>
        /// Builds the destination file name from the download URI
        /// and the destination folder (which is stored in a public property
        /// and could be set by a view that subscribes to this view model).
        /// </summary>
        /// <remarks>
        /// Derived classes will typically want to override this, as
        /// the base method uses a simple generic file name that contains
        /// the version number.
        /// </remarks>
        /// <returns>Complete path of the destination file.</returns>
        protected virtual string BuildDestinationFileName()
        {
            return System.IO.Path.Combine(
                DestinationFolder,
                String.Format("update-{0}.exe", NewVersion.ToString())
                );
        }

        /// <summary>
        /// Executes the update file. This method is called by <see cref="InstallUpdate()"/>
        /// only if the Sha1 checksum of the file meets the expectation.
        /// </summary>
        /// <remarks>
        /// The path of the downloaded file is stored in <see cref="_destinationFileName"/>.
        /// Implementations of this class may want to override this method if updating is
        /// not simply a matter of executing this file.
        /// The base method executes the file with an "/UPDATE" command line parameter.
        /// </remarks>
        protected virtual void DoInstallUpdate()
        {
            if (IsVerifiedDownload)
            {
                System.Diagnostics.Process.Start(
                    GetInstallerCommand(),
                    GetInstallerParameters()
                );
            }
        }

        /// <summary>
        /// Returns the command to execute in the shell to install the update.
        /// </summary>
        /// <returns>Command to execute.</returns>
        protected virtual string GetInstallerCommand()
        {
            return _destinationFileName;
        }

        /// <summary>
        /// Returns commandline parameters for the update installer.
        /// </summary>
        protected virtual string GetInstallerParameters()
        {
            // silencing parameters for InnoSetup installers
            return "/SP- /SILENT /SUPPRESSMSGBOXES";
        }

        #endregion

        #region Private methods

        void _client_DownloadFileCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                DownloadException = e.Error;
                IsVerifiedDownload = FileHelpers.Sha1Hash(_destinationFileName) == UpdateSha1;
                IsUpdatePending = IsVerifiedDownload;
                OnDownloadUpdateFinished();
            }
            else
            {
                DownloadException = null;
                // If the download was cancelled, remove the incomplete file from disk.
                System.IO.File.Delete(DestinationFolder);
            }
        }

        void _client_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            if (DownloadProgressChanged != null)
            {
                DownloadProgressChanged(this, e);
            }
        }

        /// <summary>
        /// Inspects the downloaded version information.
        /// </summary>
        /// <param name="sender">System.Net.WebClient instance</param>
        /// <param name="e">Event arguments</param>
        void VersionInfoClient_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            if (!e.Cancelled) {
                if (e.Error == null )
                {
                    StringReader r = new StringReader(e.Result);
                    NewVersion = new SemanticVersion(r.ReadLine());
                    DownloadUri = new Uri(r.ReadLine());
                    // Use only the first word of the line as Sha1 sum
                    // to make it compatible with the output of `sha1sum`
                    UpdateSha1 = r.ReadLine().Split(' ')[0];
                    Multiline multi = new Multiline(r.ReadToEnd(), true);
                    UpdateSummary = multi.Text;
                    IsUpdateAvailable = NewVersion > GetCurrentVersion();
                }
                else
                {
                    DownloadException = e.Error;
                    IsUpdateAvailable = false;
                }
                OnCheckForUpdateFinished();
            }
        }

        #endregion

        #region Private properties

        private string Sha1 { get; set; }

        #endregion

        #region Private fields

        private WebClient _client;
        private WebClient _versionInfoClient;
        private string _destinationFileName;

        #endregion
    }
}

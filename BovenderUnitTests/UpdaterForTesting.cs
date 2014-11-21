using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Versioning;
using System.IO;

namespace Bovender.UnitTests
{
    /// <summary>
    /// Implements the abstract Updater class for the purpose of
    /// unit testing it.
    /// </summary>
    /// <remarks>
    /// The constructor of this class writes version info data
    /// to a temporary file, so that no web access is needed.
    /// </remarks>
    class UpdaterForTesting : Updater, IDisposable
    {
        const string FAKENEWVERSION = "7.0.0";

        #region Public properties

        public string TestVersion { get; set; }

        #endregion

        #region Constructor

        public UpdaterForTesting()
            : base()
        {
            _versionInfoFile = Path.GetTempFileName();
            _fakeRemoteFile = Path.GetTempFileName();
            _fakeTargetFile = Path.GetTempFileName();
            CreateFakeUpdateFile();
            CreateVersionInfoFile();
        }

        #endregion

        #region Disposing and finalizing

        ~UpdaterForTesting()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool calledFromDisposing)
        {
            System.IO.File.Delete(_versionInfoFile);
            System.IO.File.Delete(_fakeRemoteFile);
            System.IO.File.Delete(_fakeTargetFile);
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns the URL for the XL Toolbox NG version info file.
        /// </summary>
        /// <returns>
        /// URI for http://xltoolbox.sourceforge.net/version-ng.txt
        /// </returns>
        protected override Uri GetVersionInfoUri()
        {
            return new Uri(String.Format("file:///{0}", _versionInfoFile));
        }

        /// <summary>
        /// Returns a semantic version for the version string in
        /// <see cref="TestVersion"/>.
        /// </summary>
        /// <returns>Semantic version created with <see cref="TestVersion"/></returns>
        protected override SemanticVersion GetCurrentVersion()
        {
            return new SemanticVersion(TestVersion);
        }

        protected override string BuildDestinationFileName()
        {
            return _fakeTargetFile;
        }

        /// <summary>
        /// In the testing updater, the user is always authorized
        /// to update.
        /// </summary>
        public override bool IsAuthorized
        {
            get
            {
                return true;
            }
        }

        #endregion

        #region Private methods

        private void CreateVersionInfoFile()
        {
            string s = String.Format(
                "{0}\nfile:///{1}\n{2}\nDescription of dummy download file.\n",
                FAKENEWVERSION,
                _fakeRemoteFile,
                FileHelpers.Sha1Hash(_fakeRemoteFile)
                );
            File.WriteAllText(_versionInfoFile, s);
        }

        private void CreateFakeUpdateFile()
        {
            string s = "This is a dummy update file!\n";
            File.WriteAllText(_fakeRemoteFile, s);
        }

        #endregion

        #region Private fields

        private string _versionInfoFile;
        private string _fakeRemoteFile;
        private string _fakeTargetFile;

        #endregion
    }
}

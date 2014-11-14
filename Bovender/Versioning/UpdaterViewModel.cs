using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;

namespace Bovender.Versioning
{
    /// <summary>
    /// Acts as a view model for the <see cref="Updater"/> class which is
    /// concerned with fetching version information and downloading
    /// the update. The view model implements related interactivity.
    /// </summary>
    /// <remarks>
    /// Views that subscribe to this view model should provide their own
    /// Updater class, since the base Updater class is abstract. Derived
    /// Updater classes provide project-specific information such as
    /// the download URI.
    /// </remarks>
    public class UpdaterViewModel : ViewModelBase
    {
        #region Public properties

        public Uri DownloadUri { get { return _updater.DownloadUri; } }

        public string DestinationFolder
        {
            get { return _updater.DestinationFolder; }
            set
            {
                _updater.DestinationFolder = value;
                OnPropertyChanged("DestinationFolder");
            }
        }

        #endregion

        #region Commands

        public DelegatingCommand CheckForUpdateCommand
        {
            get
            {
                if (_checkForUpdateCommand == null)
                {
                    _checkForUpdateCommand = new DelegatingCommand(
                        (param) => DoCheckForUpdate());
                }
                return _checkForUpdateCommand;
            }
        }

        public DelegatingCommand DownloadUpdateCommand
        {
            get
            {
                if (_downloadUpdateCommand == null)
                {
                    _downloadUpdateCommand = new DelegatingCommand(
                        (param) => DoDownloadUpdate()
                        );
                }
                return _downloadUpdateCommand;
            }
        }

        #endregion

        #region MVVM messages

        public Message<ViewModelMessageContent> CheckForUpdateMessage
        {
            get
            {
                if (_checkForUpdateMessage == null)
                {
                    _checkForUpdateMessage = new Message<ViewModelMessageContent>();
                }
                return _checkForUpdateMessage;
            }
        }

        public Message<ViewModelMessageContent> UpdateAvailableMessage
        {
            get
            {
                if (_updateAvailableMessage == null)
                {
                    _updateAvailableMessage = new Message<ViewModelMessageContent>();
                }
                return _updateAvailableMessage;
            }
        }

        public Message<ViewModelMessageContent> UpdateAvailableButNotAuthorizedMessage
        {
            get
            {
                if (_updateAvailableButNotAuthorizedMessage == null)
                {
                    _updateAvailableButNotAuthorizedMessage = new Message<ViewModelMessageContent>();
                }
                return _updateAvailableButNotAuthorizedMessage;
            }
        }

        public Message<ViewModelMessageContent> NoUpdateAvailableMessage
        {
            get
            {
                if (_noUpdateAvailableMessage == null)
                {
                    _noUpdateAvailableMessage = new Message<ViewModelMessageContent>();
                }
                return _noUpdateAvailableMessage;
            }
        }

        public Message<ProcessMessageContent> DownloadUpdateMessage
        {
            get
            {
                if (_downloadUpdateMessage == null)
                {
                    _downloadUpdateMessage = new Message<ProcessMessageContent>();
                }
                return _downloadUpdateMessage;
            }
        }

        public Message<ViewModelMessageContent> UpdateInstallableMessage
        {
            get
            {
                if (_updateInstallableMessage == null)
                {
                    _updateInstallableMessage = new Message<ViewModelMessageContent>();
                }
                return _updateInstallableMessage;
            }
        }

        public Message<ViewModelMessageContent> UpdateFailedVerificationMessage
        {
            get
            {
                if (_updateFailedVerificationMessage == null)
                {
                    _updateFailedVerificationMessage = new Message<ViewModelMessageContent>();
                }
                return _updateFailedVerificationMessage;
            }
        }

        public Message<ViewModelMessageContent> NetworkFailureMessage
        {
            get
            {
                if (_networkFailureMessage == null)
                {
                    _networkFailureMessage = new Message<ViewModelMessageContent>();
                }
                return _networkFailureMessage;
            }
        }

        #endregion

        #region Constructor

        public UpdaterViewModel(Updater updater) 
            : base()
        {
            _updater = updater;
        }

        #endregion

        #region Private methods

        private void DoCheckForUpdate()
        {
            CheckProcessMessageContent.CancelProcess = _updater.CancelCheckForUpdate;
            _updater.CheckForUpdateFinished += _updater_CheckForUpdateFinished;
            _updater.CheckForUpdate();
            CheckForUpdateMessage.Send(CheckProcessMessageContent);
        }

        void _updater_CheckForUpdateFinished(object sender, EventArgs e)
        {
            // Set the 'IsIndeterminate' property of the process message content to false
            // so that the progress bar stops its animation.
            CheckProcessMessageContent.IsIndeterminate = false;
            if (_updater.DownloadException == null)
            {
                if (_updater.UpdateAvailable)
                {
                    if (_updater.IsAuthorized)
                    {
                        UpdateAvailableMessage.Send(CheckProcessMessageContent);
                    }
                    else
                    {
                        UpdateAvailableButNotAuthorizedMessage.Send(CheckProcessMessageContent);
                    }
                }
                else
                {
                    NoUpdateAvailableMessage.Send(CheckProcessMessageContent);
                }
            }
            else
            {
                NetworkFailureMessage.Send(CheckProcessMessageContent);
            }
        }

        private void DoDownloadUpdate()
        {
            DownloadProcessMessageContent.CancelProcess = _updater.CancelDownload;
            _updater.DownloadProgressChanged += _updater_DownloadProgressChanged;
            _updater.DownloadUpdateFinished += _updater_DownloadUpdateFinished;
            _updater.DownloadUpdate();
            DownloadUpdateMessage.Send(DownloadProcessMessageContent);
        }

        void _updater_DownloadUpdateFinished(object sender, EventArgs e)
        {
            if (_updater.DownloadException == null)
            {
                if (_updater.IsAuthorized)
                {
                    if (_updater.IsVerifiedDownload)
                    {
                        UpdateInstallableMessage.Send(DownloadProcessMessageContent);
                    }
                    else
                    {
                        UpdateFailedVerificationMessage.Send(DownloadProcessMessageContent);
                    }
                }
                else
                {
                    UpdateAvailableButNotAuthorizedMessage.Send(DownloadProcessMessageContent);
                }
            }
            else
            {
                NetworkFailureMessage.Send(DownloadProcessMessageContent);
            }
        }

        void _updater_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            DownloadProcessMessageContent.PercentCompleted = e.ProgressPercentage;
        }

        #endregion

        #region Private Properties

        private ProcessMessageContent CheckProcessMessageContent
        {
            get
            {
                if (_checkProcessMessageContent == null)
                {
                    _checkProcessMessageContent = new ProcessMessageContent(this);
                    _checkProcessMessageContent.IsIndeterminate = true;
                }
                return _checkProcessMessageContent;
            }
        }

        private ProcessMessageContent DownloadProcessMessageContent
        {
            get
            {
                if (_downloadProcessMessageContent == null)
                {
                    _downloadProcessMessageContent = new ProcessMessageContent(this);
                }
                return _downloadProcessMessageContent;
            }
        }

        #endregion

        #region Private fields

        private Updater _updater;
        private DelegatingCommand _checkForUpdateCommand;
        private DelegatingCommand _downloadUpdateCommand;
        private Message<ViewModelMessageContent> _checkForUpdateMessage;
        private Message<ViewModelMessageContent> _updateAvailableMessage;
        private Message<ViewModelMessageContent> _updateAvailableButNotAuthorizedMessage;
        private Message<ViewModelMessageContent> _noUpdateAvailableMessage;
        private Message<ProcessMessageContent> _downloadUpdateMessage;
        private Message<ViewModelMessageContent> _updateInstallableMessage;
        private Message<ViewModelMessageContent> _updateFailedVerificationMessage;
        private Message<ViewModelMessageContent> _networkFailureMessage;
        private ProcessMessageContent _checkProcessMessageContent;
        private ProcessMessageContent _downloadProcessMessageContent;

        #endregion
    }
}

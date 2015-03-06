/* UpdaterViewModel.cs
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
    /// Acts as a view model for the <see cref="Updater"/> class which is concerned with
    /// fetching version information and downloading the update. The view model implements
    /// related interactivity.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Views that subscribe to this view model should provide their own Updater class,
    /// since the base Updater class is abstract. Derived Updater classes provide
    /// project-specific information such as the download URI.
    /// </para>
    /// <para>
    /// Checking for an available update and downloading it is a complex task that
    /// involves several interactions with a user (or a view, to stay with MVVM
    /// terminology). Basically, the process will be as follows:
    /// </para>
    /// 
    /// <list type="bullet">
    /// <item>
    /// The View subscribes to the UpdaterViewModel's <see cref="CheckForUpdateMessage"/>
    /// and executes the VM's <see cref="CheckForUpdateCommand"/>.
    /// </item>
    /// <item>
    /// The VM then arranges for the Updater model to asynchronously fetch version
    /// information from the remote repository and sends the <see cref="CheckForUpdateMessage"/>.
    /// </item>
    /// <item>
    /// Upon receiving the <see cref="CheckForUpdateMessage"/>, the View could, for example,
    /// invoke another view that displays an 'indeterminate progress'. It
    /// could however just silently wait for the update check to finish. In either
    /// case, _some_ view should listen to one of four self-explanatory messages
    /// that the VM will send upon receiving the current version information:
    ///     <list type="number">
    ///     <item><see cref="UpdateAvailableMessage"/></item>
    ///     <item><see cref="NoUpdateAvailableMessage"/></item>
    ///     <item><see cref="UpdateAvailableButNotAuthorizedMessage"/> (if the current user
    ///     does not have write permissions to the assembly's folder)</item>
    ///     <item><see cref="NetworkFailureMessage"/></item> if any exception occurred while
    ///     downloading the current version information.
    ///     </list>
    ///     </item>
    /// <item>
    /// The last three of these messages could be dealt with by simply showing
    /// an informative view to the user, for example, or writing to a log file.
    /// If the view receives the <see cref="UpdateAvailableMessage"/>, it could show the
    /// update version information that is exposed in the VM's public properties
    /// to the user.
    /// </item>
    /// <item>
    /// Before starting to download the update, the VM's <see cref="ChooseDestinationFolderCommand"/>
    /// must be executed. This will send the <see cref="ChooseDestinationFolderMessage"/>,
    /// whose content is a <see cref="StringMessageContent"/> that carries the last used
    /// destination folder (or the user's MyDocuments special folder).
    /// </item>
    /// <item>
    /// The view can now optionally have the user choose the appropriate download folder and
    /// set the VM's <see cref="DestinationFolder"/> property accordingly (e.g. by data binding).
    /// If the view is a GUI (e.g. a WPF window), this is best done by invoking a
    /// <see cref="ChooseFolderAction"/>, since there is no ready-to-use folder picker in the
    /// WPF. The ChooseFolderAction responds to the VM's ChooseDestinationFolderMessage by
    /// calling the message's respond method.
    /// </item>
    /// <item>
    /// The VM now examines the content of the ChooseDestinationFolderMessage; if the
    /// message content's Confirmed property is true, it will save the chosen folder in the
    /// assembly's properties and send the <see cref="ReadyToDownloadMessage"/>.
    /// </item>
    /// <item>
    /// The view that invoked the ChooseFolderAction listens for the ReadyToDownloadMessage
    /// and executes the VM's <see cref="DownloadUpdateCommand"/>.
    /// </item>
    /// <item>
    /// The DownloadUpdateCommand requests the Updater model to asynchronously download
    /// the update to the desired folder. It sends the <see cref="DownloadUpdateMessage"/>
    /// which contains the download progress (that can be data-bound to by a progress bar,
    /// for example) and whose CancelCommand can be invoked by a view to cancel downloading.
    /// </item>
    /// <item>
    /// The view that executed the DownloadUpdateCommand listens for a DownloadUpdateMessage
    /// and could invoke a <see cref="ShowViewAction"/> with the message argument that shows a
    /// progress bar view which is data-bound to the DownloadUpdateMessage's ProcessMessageContent.
    /// </item>
    /// <item>
    /// When the download is finished, the VM compares the file's Sha1 checksum with the
    /// expected checksum that was announced in the update version information and issues an
    /// <see cref="UpdateInstallableMessage"/> or a <see cref="UpdateFailedVerificationMessage"/>
    /// as appropriate. It may also send a <see cref="NetworkFailureMessage"/> if some exception
    /// occurred while downloading.
    /// </item>
    /// <item>
    /// Finally, upon receiving the UpdateInstallableMessage, the view may execute the VM's
    /// <see cref="InstallUpdateCommand"/>.
    /// </item>
    /// The VM's InstallUpdateCommand logic will proceed with executing the downloaded file
    /// (by calling the updater model's InstallUpdate method), or send an UpdateFailedVerificationMessage.
    /// </list>
    /// <para>
    /// Note: Views (or code) that subscribe to the UpdaterViewModel do not necessarily have
    /// to listen and respond to all of the VM's messages.
    /// For example, it is not necessary to execute the ChooseDestinationFolderCommand, listen for
    /// the ChooseDestinationFolderMessage, respond to it, and listen for the ReadyToDownloadMessage.
    /// These commands and messages serve to facilitate displaying GUI views in the MVVM pattern.
    /// If the code that wishes to check for an available update (e.g., in the background without
    /// the user's attention) has another way to set the destination folder or is happy with the
    /// last used or default destination folder, it may just proceed with executing the
    /// DownloadUpdate command.
    /// </para>
    /// </remarks>
    public class UpdaterViewModel : ViewModelBase
    {
        #region Public properties

        public SemanticVersion NewVersion { get { return _updater.NewVersion; } }

        public SemanticVersion CurrentVersion { get { return _updater.CurrentVersion;  } }

        public string UpdateSummary { get { return _updater.UpdateSummary; } }

        public Uri DownloadUri { get { return _updater.DownloadUri; } }

        public bool IsUserAuthorized { get { return _updater.IsAuthorized; } }

        public bool IsVerifiedDownload { get { return _updater.IsVerifiedDownload; } }

        public bool IsUpdatePending { get { return _updater.IsUpdatePending; } }

        public bool CanCheckForUpdate
        {
            get
            {
                return !IsLocked && !IsUpdatePending;
            }
        }

        public string DestinationFolder
        {
            get
            {
                if (String.IsNullOrEmpty(_updater.DestinationFolder))
                {
                    string s = Properties.Settings.Default.UpdateDestinationFolder;
                    if (string.IsNullOrEmpty(s))
                    {
                        s = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    };
                    _updater.DestinationFolder = s;
                    OnPropertyChanged("DestinationFolder");
                };
                return _updater.DestinationFolder;
            }

            set
            {
                _updater.DestinationFolder = value;
                OnPropertyChanged("DestinationFolder");
            }
        }

        public Exception DownloadException { get { return _updater.DownloadException;  } }

        #endregion

        #region Commands

        public DelegatingCommand CheckForUpdateCommand
        {
            get
            {
                if (_checkForUpdateCommand == null)
                {
                    _checkForUpdateCommand = new DelegatingCommand(
                        (param) => DoCheckForUpdate(),
                        (param) => CanCheckForUpdate);
                }
                return _checkForUpdateCommand;
            }
        }

        public DelegatingCommand ChooseDestinationFolderCommand
        {
            get
            {
                if (_chooseDestinationFolderCommand == null)
                {
                    _chooseDestinationFolderCommand = new DelegatingCommand(
                        (param) => DoChooseDestinationFolder(),
                        (param) => CanChooseDestinationFolder());
                }
                return _chooseDestinationFolderCommand;
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

        public DelegatingCommand InstallUpdateCommand
        {
            get
            {
                if (_installUpdateCommand == null)
                {
                    _installUpdateCommand = new DelegatingCommand(
                        (param) => DoInstallUpdate(),
                        (param) => CanInstallUpdate());
                }
                return _installUpdateCommand;
            }
        }

        #endregion

        #region MVVM messages

        public Message<ProcessMessageContent> CheckForUpdateMessage
        {
            get
            {
                if (_checkForUpdateMessage == null)
                {
                    _checkForUpdateMessage = new Message<ProcessMessageContent>();
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
            private set
            {
                _updateAvailableMessage = value;
            }
        }

        public Message<ProcessMessageContent> ReadyToDownloadMessage
        {
            get
            {
                if (_readyToDownloadMessage == null)
                {
                    _readyToDownloadMessage = new Message<ProcessMessageContent>();
                }
                return _readyToDownloadMessage;
            }
            private set
            {
                _readyToDownloadMessage = value;
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

        public Message<StringMessageContent> ChooseDestinationFolderMessage
        {
            get
            {
                if (_chooseDestinationFolderMessage == null)
                {
                    _chooseDestinationFolderMessage = new Message<StringMessageContent>();
                }
                return _chooseDestinationFolderMessage;
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
            CheckProcessMessageContent.CancelProcess = CancelCheckForUpdate;
            DownloadProcessMessageContent.CancelProcess = CancelDownloadUpdate;
            _updater.CheckForUpdateFinished += _updater_CheckForUpdateFinished;
            _updater.DownloadProgressChanged += _updater_DownloadProgressChanged;
            _updater.DownloadUpdateFinished += _updater_DownloadUpdateFinished;
        }

        #endregion

        #region Private methods

        private void DoCheckForUpdate()
        {
            IsLocked = true;
            CheckForUpdateMessage.Send(CheckProcessMessageContent);
            _updater.CheckForUpdate();
        }

        void _updater_CheckForUpdateFinished(object sender, EventArgs e)
        {
            IsLocked = false;
            Action action = new Action(() =>
                { 
                    // Set the 'IsIndeterminate' property of the process message content to false
                    // so that the progress bar stops its animation.
                    CheckProcessMessageContent.IsIndeterminate = false;
                    if (_updater.DownloadException == null)
                    {
                        if (_updater.IsUpdateAvailable)
                        {
                            if (_updater.IsAuthorized)
                            {
                                UpdateAvailableMessage.Send(
                                    CheckProcessMessageContent,
                                    (content) => InstallUpdateCommand.Execute(null));
                            }
                            else
                            {
                                UpdateAvailableButNotAuthorizedMessage.Send(
                                    CheckProcessMessageContent,
                                    (content) => CloseViewCommand.Execute(null));
                            }
                        }
                        else
                        {
                            NoUpdateAvailableMessage.Send(
                                CheckProcessMessageContent,
                                (content) => CloseViewCommand.Execute(null));
                        }
                    }
                    else
                    {
                        OnPropertyChanged("DownloadException");
                        NetworkFailureMessage.Send(
                            CheckProcessMessageContent,
                            (content) => CloseViewCommand.Execute(null));
                    }
                });

            // Asynchronous operations that interact with GUI views must be invoked
            // via the view's dispatcher. If there is no GUI view and hence no
            // dispatcher, simply invoke the action itself (e.g., if the view
            // is a non-GUI object).
            if (ViewDispatcher != null)
            {
                ViewDispatcher.Invoke(action);
            }
            else {
                action.Invoke();
            }
        }

        private void DoChooseDestinationFolder()
        {
            ChooseDestinationFolderMessage.Send(
                new StringMessageContent(DestinationFolder),
                (StringMessageContent returnContent) =>
                {
                    if (returnContent.Confirmed)
                    {
                        DestinationFolder = returnContent.Value;
                        Properties.Settings.Default.UpdateDestinationFolder = DestinationFolder;
                        Properties.Settings.Default.Save();
                        ReadyToDownloadMessage.Send();
                    };
                }
            );
        }

        private bool CanChooseDestinationFolder()
        {
            return _updater.IsAuthorized;
        }

        private void DoDownloadUpdate()
        {
            IsLocked = true;
            DownloadUpdateMessage.Send(DownloadProcessMessageContent);
            _updater.DownloadUpdate();
        }

        void _updater_DownloadUpdateFinished(object sender, EventArgs e)
        {
            IsLocked = false;
            // When the download process is finished, the Updater model will verify
            // the downloaded file's Sha1. Since the view model exposes the
            // updater's IsVerifiedDownload, we need to notify subscribers to the
            // view model's INotifyPropertyChanged interface (inherited from
            // ViewModelBase).
            OnPropertyChanged("IsVerifiedDownload");
            OnPropertyChanged("IsUpdatePending");

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

        private void DoInstallUpdate()
        {
            if (CanInstallUpdate())
            {
                DoCloseView();
                _updater.InstallUpdate();
            }
            else
            {
                throw new InvalidOperationException("Cannot install update: User not authorized or update not verified.");
            }
        }

        private bool CanInstallUpdate()
        {
            return (_updater.IsAuthorized && _updater.IsVerifiedDownload);
        }

        #endregion

        #region Private Properties

        private bool IsLocked
        {
            get
            {
                return _isLocked;
            }
            set
            {
                _isLocked = value;
                OnPropertyChanged("CanCheckForUpdate");
            }
        }

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

        private void CancelCheckForUpdate()
        {
            IsLocked = false;
            _updater.CancelCheckForUpdate();
        }

        private void CancelDownloadUpdate()
        {
            IsLocked = false;
             _updater.CancelDownload();
        }

        #endregion

        #region Private fields

        private Updater _updater;
        private DelegatingCommand _checkForUpdateCommand;
        private DelegatingCommand _downloadUpdateCommand;
        private DelegatingCommand _installUpdateCommand;
        private DelegatingCommand _chooseDestinationFolderCommand;
        private Message<ProcessMessageContent> _checkForUpdateMessage;
        private Message<ProcessMessageContent> _readyToDownloadMessage;
        private Message<ProcessMessageContent> _downloadUpdateMessage;
        private Message<ViewModelMessageContent> _updateAvailableMessage;
        private Message<ViewModelMessageContent> _updateAvailableButNotAuthorizedMessage;
        private Message<ViewModelMessageContent> _noUpdateAvailableMessage;
        private Message<ViewModelMessageContent> _updateInstallableMessage;
        private Message<ViewModelMessageContent> _updateFailedVerificationMessage;
        private Message<ViewModelMessageContent> _networkFailureMessage;
        private Message<StringMessageContent> _chooseDestinationFolderMessage;
        private ProcessMessageContent _checkProcessMessageContent;
        private ProcessMessageContent _downloadProcessMessageContent;
        private bool _isLocked;

        #endregion

        #region Implementation of ViewModelBase's abstract methods

        public override object RevealModelObject()
        {
            return _updater;
        }

        #endregion
    }
}

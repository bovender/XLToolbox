using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Versioning;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.About
{
    public class AboutViewModel : ViewModelBase
    {
        #region Public properties

        public SemanticVersion Version
        {
            get
            {
                return XLToolbox.Versioning.SemanticVersion.CurrentVersion();
            }
        }

        #endregion

        #region Commands

        public DelegatingCommand ShowWebsiteCommand
        {
            get
            {
                if (_showWebsiteCommand == null) {
                    _showWebsiteCommand = new DelegatingCommand(
                        (param) =>
                        {
                            System.Diagnostics.Process.Start(Properties.Settings.Default.WebsiteUrl);
                            DoCloseView();
                        },
                        null
                        );
                };
                return _showWebsiteCommand;
            }
        }

        public DelegatingCommand ShowLicenseCommand
        {
            get
            {
                if (_showLicenseCommand == null)
                {
                    _showLicenseCommand = new DelegatingCommand(
                        (param) => ShowLicenseMessage.Send()
                        );
                }
                return _showLicenseCommand;
            }
        }

        public DelegatingCommand ShowCreditsCommand
        {
            get
            {
                if (_showCreditsCommand == null)
                {
                    _showCreditsCommand = new DelegatingCommand(
                        (param) => ShowCreditsMessage.Send()
                        );
                }
                return _showCreditsCommand;
            }
        }

        #endregion

        #region MVVM messaging events

        public Message<MessageContent> ShowLicenseMessage
        {
            get
            {
                if (_showLicenseMessage == null)
                {
                    _showLicenseMessage = new Message<MessageContent>();
                }
                return _showLicenseMessage;
            }
        }

        public Message<MessageContent> ShowCreditsMessage
        {
            get
            {
                if (_showCreditsMessage == null)
                {
                    _showCreditsMessage = new Message<MessageContent>();
                }
                return _showCreditsMessage;
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Response action for the <see cref="ShowWebsiteMessage"/> message.
        /// </summary>
        /// <param name="messageContent"></param>
        private void WebsiteMessageResponse(StringMessageContent messageContent)
        {
            DoCloseView();
        }

        #endregion

        #region Private fields

        private DelegatingCommand _showWebsiteCommand;
        private DelegatingCommand _showLicenseCommand;
        private DelegatingCommand _showCreditsCommand;
        private Message<MessageContent> _showLicenseMessage;
        private Message<MessageContent> _showCreditsMessage;

        #endregion

    }
}

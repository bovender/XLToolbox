using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Versioning;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;

namespace XLToolbox.Mvvm.ViewModels
{
    public class AboutViewModel : ViewModelBase
    {
        #region Public properties

        public SemanticVersion Version
        {
            get
            {
                return SemanticVersion.CurrentVersion();
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
                        (param) => ShowWebsiteMessage.Send() 
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

        public Message<MessageContent> ShowWebsiteMessage
        {
            get
            {
                if (_showWebsiteMessage == null)
                {
                    _showWebsiteMessage = new Message<StringMessageContent>();
                }
                return _showWebsiteMessage;
            }
        }

        public Message<MessageContent> ShowLicenseMessage
        {
            get
            {
                if (_showLicenseMessage == null)
                {
                    _showLicenseMessage = new Message<StringMessageContent>();
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
                    _showCreditsMessage = new Message<StringMessageContent>();
                }
                return _showCreditsMessage;
            }
        }

        #endregion

        #region Private fields

        private DelegatingCommand _showWebsiteCommand;
        private DelegatingCommand _showLicenseCommand;
        private DelegatingCommand _showCreditsCommand;
        private Message<MessageContent> _showWebsiteMessage;
        private Message<MessageContent> _showLicenseMessage;
        private Message<MessageContent> _showCreditsMessage;

        #endregion
    }
}

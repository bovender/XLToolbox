using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Mvvm.ViewModels;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm;
using Bovender.Versioning;

namespace XLToolbox.Greeter
{
    /// <summary>
    /// View model for the greeter screen.
    /// </summary>
    public class GreeterViewModel : ViewModelBase
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

        public DelegatingCommand WhatsNewCommand
        {
            get
            {
                if (_whatsNewCommand == null)
                {
                    _whatsNewCommand = new DelegatingCommand(
                        (param) => DoShowWhatsNew()
                        );
                }
                return _whatsNewCommand;
            }
        }

        public DelegatingCommand DonateCommand
        {
            get
            {
                if (_donateCommand == null)
                {
                    _donateCommand = new DelegatingCommand(
                        (param) => DoShowDonatePage()
                        );
                }
                return _donateCommand;
            }
        }

        #endregion

        /*
        #region MVVM messages

        public Message<MessageContent> WhatsNewMessage
        {
            get
            {
                if (_whatsNewMessage == null)
                {
                    _whatsNewMessage = new Message<MessageContent>();
                }
                return _whatsNewMessage;
            }
        }

        public Message<MessageContent> DonateMessage
        {
            get
            {
                if (_donateMessage == null)
                {
                    _donateMessage = new Message<MessageContent>();
                }
                return _donateMessage;
            }
        }

        #endregion
        */

        #region Private methods

        private void DoShowWhatsNew()
        {
            System.Diagnostics.Process.Start(Properties.Settings.Default.WhatsNewUrl);
            CloseViewCommand.Execute(null);
        }

        private void DoShowDonatePage()
        {
            System.Diagnostics.Process.Start(Properties.Settings.Default.DonateUrl);
            CloseViewCommand.Execute(null);
        }

        #endregion

        #region Private fields

        private DelegatingCommand _whatsNewCommand;
        private DelegatingCommand _donateCommand;

        /*
        private Message<MessageContent> _whatsNewMessage;
        private Message<MessageContent> _donateMessage;
         */

        #endregion

        #region Overrides

        public override object RevealModelObject()
        {
            return null;
        }

        #endregion
    }
}

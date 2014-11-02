using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Bovender.Mvvm;

namespace Bovender.Mvvm.Messaging
{
    /// <summary>
    /// Holds information about percent completion of a process
    /// and defines events that occur when the process is finished.
    /// </summary>
    public class ProcessMessageContent : MessageContent, INotifyPropertyChanged
    {
        #region Public properties

        public double PercentCompleted
        {
            get
            {
                return _percentCompleted;
            }
            set
            {
                _percentCompleted = value;
                OnPropertyChanged("PercentCompleted");
            }
        }
        public bool WasSuccessful { get; protected set; }
        public bool WasCancelled { get; protected set; }

        /// <summary>
        /// Delegate that can be called to cancel the current process.
        /// </summary>
        public Action CancelProcess { get; set; }

        #endregion

        #region Commands

        public DelegatingCommand CancelCommand
        {
            get
            {
                if (_cancelCommand == null)
                {
                    _cancelCommand = new DelegatingCommand(
                        (param) => DoCancel(),
                        (param) => CanCancel()
                        );
                }
                return _cancelCommand;
            }
        }

        #endregion

        #region MVVM messages

        public Message<ProcessMessageContent> CompletedMessage
        {
            get
            {
                if (_completedMessage == null)
                {
                    _completedMessage = new Message<ProcessMessageContent>();
                }
                return _completedMessage;
            }
        }

        #endregion

        #region Constructors

        public ProcessMessageContent() : base() { }

        public ProcessMessageContent(Action cancelProcess)
            : this()
        {
            CancelProcess = cancelProcess;
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Disables the Confirm command if the Value is invalid.
        /// </summary>
        /// <returns>True if the Value is valid and the dialog can be closed.</returns>
        protected override bool CanConfirm()
        {
            return String.IsNullOrEmpty((this as IDataErrorInfo)["Value"]);
        }

        #endregion

        #region Protected methods

        /// <summary>
        /// Cancels the process, sends a message to any subscribed
        /// views informing about the state change, and requests to
        /// close the subscribed views.
        /// </summary>
        protected void DoCancel()
        {
            if (CanCancel())
            {
                CancelProcess.Invoke();
                WasCancelled = true;
                WasSuccessful = false;
                CompletedMessage.Send(this, null);
                DoCloseView();
            }
        }

        protected bool CanCancel()
        {
            return (CancelProcess != null);
        }

        #endregion

        #region Private fields

        private double _percentCompleted;
        private DelegatingCommand _cancelCommand;
        private Message<ProcessMessageContent> _completedMessage;

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Bovender.Mvvm;
using Bovender.Mvvm.ViewModels;

namespace Bovender.Mvvm.Messaging
{
    /// <summary>
    /// Holds information about percent completion of a process
    /// and defines events that occur when the process is finished.
    /// </summary>
    public class ProcessMessageContent : ViewModelMessageContent
    {
        #region Public properties

        public bool Processing
        {
            get { return _processing; }
            set
            {
                _processing = value;
                OnPropertyChanged("Processing");
            }
        }

        public bool IsIndeterminate
        {
            get { return _isIndeterminate; }
            set
            {
                _isIndeterminate = value;
                OnPropertyChanged("IsInfinite");
            }
        }

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
        
        public bool WasSuccessful
        {
            get { return _wasSuccessful; }
            set
            {
                _wasSuccessful = value;
                OnPropertyChanged("WasSuccessful");
            }
        }

        public bool WasCancelled
        {
            get { return _wasCancelled; }
            set
            {
                _wasCancelled = value;
                OnPropertyChanged("WasCancelled");
            }
        }

        /// <summary>
        /// If something in the process went wrong, this will be the corresponding
        /// exception.
        /// </summary>
        public Exception Exception
        {
            get { return _exception; }
            set
            {
                _exception = value;
                OnPropertyChanged("Exception");
            }
        }

        /// <summary>
        /// Delegate that can be called to cancel the current process.
        /// </summary>
        public Action CancelProcess
        {
            get
            {
                return _cancelProcess;
            }
            set
            {
                _cancelProcess = value;
                OnPropertyChanged("CancelProcess");
            }
        }

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

        public ProcessMessageContent(ViewModelBase viewModel) : base(viewModel) { }

        public ProcessMessageContent(Action cancelProcess)
            : this()
        {
            CancelProcess = cancelProcess;
        }

        public ProcessMessageContent(ViewModelBase viewModel, Action cancelProcess)
            : base(viewModel)
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
        private bool _processing;
        private bool _wasCancelled;
        private bool _wasSuccessful;
        private Exception _exception;
        private Action _cancelProcess;
        private bool _isIndeterminate;

        #endregion
    }
}

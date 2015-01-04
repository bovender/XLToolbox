/* ProcessMessageContent.cs
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
using Bovender.Mvvm.ViewModels;
using System;
using System.ComponentModel;

namespace Bovender.Mvvm.Messaging
{
    /// <summary>
    /// Holds information about percent completion of a process
    /// and defines events that occur when the process is finished.
    /// This message can optionally carry a view model with additional
    /// events and capabilities.
    /// </summary>
    public class ProcessMessageContent : ViewModelMessageContent
    {
        #region Public properties

        public string Caption { get; set; }

        public string Message { get; set; }

        public string CancelButtonText { get; set; }

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
                OnPropertyChanged("IsIndeterminate");
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

        /// <summary>
        /// Creates a new ProcessMessageContent.
        /// </summary>
        public ProcessMessageContent() : base() { }

        /// <summary>
        /// Creates a new ProcessMessageContent that encapsulates
        /// the given <paramref name="viewModel"/> to enable views
        /// to access the encapsulated view model's members
        /// (e.g. in order to bind to the view model's events).
        /// </summary>
        /// <param name="viewModel">View model to encapsulate</param>
        public ProcessMessageContent(ViewModelBase viewModel) : base(viewModel) { }

        /// <summary>
        /// Creates a new ProcessMessageContent that has the ability
        /// to cancel a process.
        /// </summary>
        /// <param name="cancelProcess">Method to invoke when the Cancel button
        /// is clicked</param>
        public ProcessMessageContent(Action cancelProcess)
            : this()
        {
            CancelProcess = cancelProcess;
        }

        /// <summary>
        /// Creates a new ProcessMessageContent that has the ability
        /// to cancel a process and that encapsulates a view model to
        /// provide a view easy access to the view model's members
        /// (e.g. in order to bind to the view model's events).
        /// </summary>
        /// <param name="viewModel">View model to encapsulate</param>
        /// <param name="cancelProcess">Method to invoke when the Cancel button
        /// is clicked</param>
        public ProcessMessageContent(ViewModelBase viewModel, Action cancelProcess)
            : base(viewModel)
        {
            CancelProcess = cancelProcess;
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

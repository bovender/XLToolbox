/* ExceptionViewModel.cs
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
using System.Collections.Specialized;
using Bovender.Mvvm;
using Bovender.Mvvm.Messaging;
using Bovender.Mvvm.ViewModels;

namespace Bovender.ExceptionHandler
{
    /// <summary>
    /// Provides easy access to several system properties that are
    /// relevant for bug reports. 
    /// </summary>
    public abstract class ExceptionViewModel : ViewModelBase
    {
        #region Public properties

        public string User
        {
            get { return _user; }
            set
            {
                _user = value;
                OnPropertyChanged("User");
            }
        }

        public string Email
        {
            get { return _email; }
            set
            {
                _email = value;
                OnPropertyChanged("Email");
                OnPropertyChanged("IsCcUserEnabled");
            }
        }

        public bool CcUser
        {
            get { return _ccUser; }
            set
            {
                _ccUser = value;
                OnPropertyChanged("CcUser");
            }
        }

        public bool IsCcUserEnabled
        {
            get
            {
                // TODO: Check if it is really an e-mail address
                        return !String.IsNullOrEmpty(Email);
            }
        }

        public string Comment
        {
            get
            {
                return _comment;
            }
            set
            {
                _comment = value;
                OnPropertyChanged("Comment");
            }
        }

        public string Exception { get; private set; }
        public string Message { get; private set; }
        public string InnerException { get; private set; }
        public string InnerMessage { get; private set; }

        public bool HasInnerException
        {
            get
            {
                return !String.IsNullOrEmpty(InnerException);
            }
        }

        public string StackTrace { get; private set; }

        public string OS
        {
            get
            {
                return Environment.OSVersion.VersionString;
            }
        }

        public string CLR
        {
            get
            {
                return Environment.Version.ToString();
            }
        }

        public string ProcessBitness
        {
            get
            {
                return Environment.Is64BitProcess ? "64-bit" : "32-bit";
            }
        }

        public string OSBitness
        {
            get
            {
                return Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit";
            }
        }

        public string ReportID { get; private set; }

        #endregion

        #region Commands

        public DelegatingCommand SubmitReportCommand
        {
            get
            {
                if (_submitReportCommand == null)
                {
                    _submitReportCommand = new DelegatingCommand(
                        (param) => DoSubmitReport(),
                        (param) => CanSubmitReport()
                        );
                }
                return _submitReportCommand;
            }
        }

        public DelegatingCommand ViewDetailsCommand
        {
            get
            {
                if (_viewDetailsCommand == null)
                {
                    _viewDetailsCommand = new DelegatingCommand(
                        (param) => ViewDetailsMessage.Send(
                            new ViewModelMessageContent(this),
                            null)
                        );
                }
                return _viewDetailsCommand;
            }
        }

        public DelegatingCommand ClearFormCommand
        {
            get {
            if (_clearFormCommand == null) {
                _clearFormCommand = new DelegatingCommand(
                    (param) => DoClearForm(),
                    (param) => CanClearForm()
                    );
            }
            return _clearFormCommand;
            }

        }
        #endregion

        #region MVVM messages

        /// <summary>
        /// Signals that more details about the exception are requested to be shown.
        /// </summary>
        public Message<ViewModelMessageContent> ViewDetailsMessage
        {
            get
            {
                if (_viewDetailsMessage == null)
                {
                    _viewDetailsMessage = new Message<ViewModelMessageContent>();
                }
                return _viewDetailsMessage;
            }
        }

        /// <summary>
        /// Signals that an exception report is being posted to the online
        /// issue tracker.
        /// </summary>
        public Message<MessageContent> SubmitReportMessage
        {
            get
            {
                if (_submitReportMessage == null)
                {
                    _submitReportMessage = new Message<MessageContent>();
                }
                return _submitReportMessage;
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Instantiates the class and sets the report ID to the hexadecimal
        /// representation of the current ticks (time elapsed since 1 AD).
        /// </summary>
        public ExceptionViewModel(Exception e)
        {
            /* To produce a 'unique' error ID, we take the system time in ticks
             * elapsed since 1/1/2010, bit-shift it by 20 bits (empirically determined
             * by balancing resolution with capacity of this code), then
             * converting it to a hexadecimal string represenation.
             */
            long baseDate = (new DateTime(2010, 1, 1)).Ticks >> 20;
            long now = DateTime.Now.Ticks >> 20;
            ReportID = Convert.ToString(now - baseDate, 16);

            string devPath = DevPath();
            Exception = e.ToString().Replace(devPath, String.Empty);
            Message = e.Message;
            if (e.InnerException != null)
            {
                InnerException = e.InnerException.ToString().Replace(devPath, String.Empty);
                InnerMessage = e.InnerException.Message;
            }
            else
            {
                InnerException = "";
                InnerMessage = "";
            }
            StackTrace = e.StackTrace.Replace(devPath, String.Empty);

            User = Settings.User;
            Email = Settings.Email;
            CcUser = Settings.CcUser;
        }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Helper methods that returns a URI to POST the exception report to.
        /// </summary>
        /// <returns>Valid URI of a server that accepts POST requests.</returns>
        protected abstract Uri GetPostUri();

        #endregion

        #region Overrides

        protected override void DoCloseView()
        {
            Settings.User = User;
            Settings.Email = Email;
            Settings.CcUser = CcUser;
            Settings.Save();
            base.DoCloseView();
        }

        #endregion

        #region Private methods

        private void webClient_UploadValuesCompleted(object sender, UploadValuesCompletedEventArgs e)
        {
            // Set 'IsIndeterminate' to false to stop the ProgressBar animation.
            SubmissionProcessMessageContent.IsIndeterminate = false;
            SubmissionProcessMessageContent.WasSuccessful = false;
            if (!e.Cancelled)
            {
                SubmissionProcessMessageContent.WasCancelled = false;
                if (e.Error == null)
                {
                    string result = Encoding.UTF8.GetString(e.Result);
                    if (result == ReportID)
                    {
                        SubmissionProcessMessageContent.WasSuccessful = true;
                    }
                    else
                    {
                        SubmissionProcessMessageContent.Exception = new UnexpectedResponseException(
                            String.Format(
                                "Received an unexpected return value from the web service (should be report ID {0}).",
                                ReportID
                            )
                        );
                    }
                }
                else
                {
                    SubmissionProcessMessageContent.Exception = e.Error;
                }
            }
            else
            {
                SubmissionProcessMessageContent.WasCancelled = true;
            }
            SubmissionProcessMessageContent.Processing = false;
            // Notify any subscribed views that the process is completed.
            SubmissionProcessMessageContent.CompletedMessage.Send(SubmissionProcessMessageContent);
        }

        private void CancelSubmission()
        {
            if (_webClient != null)
            {
                _webClient.CancelAsync();
            }
        }

        #endregion

        #region Protected methods

        protected virtual void DoSubmitReport()
        {
            SubmissionProcessMessageContent.CancelProcess = new Action(CancelSubmission);
            SubmissionProcessMessageContent.Processing = true;
            _webClient = new WebClient();
            NameValueCollection v = GetPostValues();
            _webClient.UploadValuesCompleted += webClient_UploadValuesCompleted;
            _webClient.UploadValuesAsync(GetPostUri(), v);
            SubmitReportMessage.Send(SubmissionProcessMessageContent);
        }

        protected virtual bool CanSubmitReport()
        {
            return ((GetPostUri() != null) && !SubmissionProcessMessageContent.Processing);
        }

        protected virtual void DoClearForm()
        {
            User = String.Empty;
            Email = String.Empty;
            Comment = String.Empty;
            CcUser = true;
        }

        protected virtual bool CanClearForm()
        {
            return !(
                String.IsNullOrEmpty(User) &&
                String.IsNullOrEmpty(Email) &&
                String.IsNullOrEmpty(Comment)
                );
        }

        /// <summary>
        /// Returns a collection of key-value pairs of exception context information
        /// that will be submitted to the exception reporting server.
        /// </summary>
        /// <returns>Collection of key-value pairs with exception context information</returns>
        protected virtual NameValueCollection GetPostValues()
        {
            NameValueCollection v = new NameValueCollection(20);
            v["report_id"] = ReportID;
            v["usersName"] = User;
            v["usersMail"] = Email;
            v["ccUser"] = CcUser.ToString();
            v["exception"] = Exception;
            v["message"] = Message;
            v["comment"] = Comment;
            v["inner_exception"] = InnerException;
            v["inner_message"] = InnerMessage;
            v["stack_trace"] = StackTrace;
            v["process_bitness"] = ProcessBitness;
            v["operating_system"] = OS;
            v["os_bitness"] = OSBitness;
            v["clr_version"] = CLR;
            return v;
        }

        /// <summary>
        /// Returns the path on the development machine that shall be stripped
        /// from the file information in the exception and stack trace.
        /// </summary>
        /// <remarks>
        /// If an application is distributed along with .pdb files, the entire path of
        /// files on the development machine is included in an exception message. Since
        /// pdb files are required in order to get the line on which an exception occurred,
        /// this method provides a way to define which part of the path shall be removed.
        /// </remarks>
        /// <example>
        /// x:\XLToolbox\NG
        /// </example>
        /// <returns>String.Empty by default; derived classes should override this.</returns>
        protected virtual string DevPath()
        {
            return String.Empty;
        }

        #endregion

        #region Protected properties

        protected ProcessMessageContent SubmissionProcessMessageContent
        {
            get
            {
                if (_submissionProcessMessageContent == null)
                {
                    _submissionProcessMessageContent = new ProcessMessageContent(
                        new Action(CancelSubmission)
                        );
                    _submissionProcessMessageContent.IsIndeterminate = true;
                }
                return _submissionProcessMessageContent;
            }
        }

        #endregion

        #region Private fields

        private string _user;
        private string _email;
        private string _comment;
        private bool _ccUser;
        private WebClient _webClient;
        private DelegatingCommand _submitReportCommand;
        private DelegatingCommand _viewDetailsCommand;
        private DelegatingCommand _clearFormCommand;
        private Message<MessageContent> _submitReportMessage;
        private Message<ViewModelMessageContent> _viewDetailsMessage;
        private ProcessMessageContent _submissionProcessMessageContent;

        #endregion
    }
}

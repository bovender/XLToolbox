using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Collections.Specialized;
using Bovender.Mvvm.ViewModels;

namespace Bovender.ExceptionHandler
{
    /// <summary>
    /// Provides easy access to several system properties that are
    /// relevant for bug reports. 
    /// </summary>
    public abstract class ExceptionViewModel : ViewModelBase
    {
        #region Events

        public event EventHandler<UploadValuesCompletedEventArgs> UploadSuccessful;
        public event EventHandler<UploadFailedEventArgs> UploadFailed;

        #endregion

        #region Public properties

        public string User { get; set; }
        public string Email { get; set; }
        public bool CcUser { get; set; }
        public string Comment { get; set; }

        public string Exception { get; private set; }
        public string Message { get; private set; }
        public string InnerException { get; private set; }
        public string InnerMessage { get; private set; }
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

            Exception = e.ToString();
            Message = e.Message;
            if (e.InnerException != null)
            {
                InnerException = e.InnerException.ToString();
                InnerMessage = e.InnerException.Message;
            }
            else
            {
                InnerException = "";
                InnerMessage = "";
            }
            StackTrace = e.StackTrace;
        }

        #endregion

        #region Abstract methods

        /// <summary>
        /// Helper methods that returns a URI to POST the exception report to.
        /// </summary>
        /// <returns>Valid URI of a server that accepts POST requests.</returns>
        protected abstract Uri GetPostUri();

        #endregion

        #region Private methods

        /// <summary>
        /// Sends a POST request with the exception data to the web service
        /// at http://xltoolbox.sourceforge.net/receive.php
        /// </summary>
        public void Send()
        {
            using (WebClient client = new WebClient())
            {
                NameValueCollection v = GetPostValues();
                client.UploadValuesCompleted += client_UploadValuesCompleted;
                client.UploadValuesAsync(GetPostUri(), v);
            }
        }

        private void client_UploadValuesCompleted(object sender, UploadValuesCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                string result = Encoding.UTF8.GetString(e.Result);
                if (result == ReportID)
                {
                    OnUploadSuccessful(e);
                }
                else
                {
                    /* To signal that the return value has unexpected content,
                     * we add a new instance of an exception to the UploadValuesCompletedEventArgs
                     * variable.
                     */
                    UnexpectedResponseException error = new UnexpectedResponseException(
                        "Received an unexpected response from the web service.");
                    UploadFailedEventArgs args = new UploadFailedEventArgs(error);
                    OnUploadFailed(args);
                }
            } else
	        {
                UploadFailedEventArgs args = new UploadFailedEventArgs(e.Error);
                OnUploadFailed(args);
	        }
        }

        #endregion

        #region Protected methods

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

        protected virtual void OnUploadSuccessful(UploadValuesCompletedEventArgs e)
        {
            if (UploadSuccessful != null)
            {
                UploadSuccessful(this, e);
            }
        }

        protected virtual void OnUploadFailed(UploadFailedEventArgs e)
        {
            if (UploadFailed != null)
            {
                UploadFailed(this, e);
            }
        }

        #endregion
    }
}

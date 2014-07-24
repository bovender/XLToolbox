using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Collections.Specialized;
using XLToolbox.Version;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Error
{
    /// <summary>
    /// Provides easy access to several system properties that are
    /// relevant for bug reports. 
    /// </summary>
    public class Reporter
    {
        public event EventHandler<UploadValuesCompletedEventArgs> UploadSuccessful;
        public event EventHandler<UploadFailedEventArgs> UploadFailed;

        private const string _postUrl = "http://xltoolbox.sourceforge.net/receive.php";
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

        public string ExcelVersion { get; set; }

        public string ExcelBitness
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

        public string ToolboxVersion
        {
            get
            {
                return SemanticVersion.CurrentVersion().ToString();
            }
        }

        public string FreeImageVersion
        {
            get
            {
                if (FreeImageAPI.FreeImage.IsAvailable())
                {
                    return FreeImageAPI.FreeImage.GetVersion();
                }
                else
                {
                    return "Not available";
                }
            }
        }

        public string ReportID { get; private set; }

        /// <summary>
        /// Instantiates the class and sets the report ID to the hexadecimal
        /// representation of the current ticks (time elapsed since 1 AD).
        /// </summary>
        public Reporter(Application a, Exception e)
        {
            /* To produce a 'unique' error ID, we take the system time in ticks
             * elapsed since 1/1/2010, bit-shift it by 20 bits (empirically determined
             * by balancing resolution with capacity of this code), then
             * converting it to a hexadecimal string represenation.
             */
            long baseDate = (new DateTime(2010, 1, 1)).Ticks >> 20;
            long now = DateTime.Now.Ticks >> 20;
            ReportID = Convert.ToString(now - baseDate, 16);

            ExcelVersion = a.Version;
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

        /// <summary>
        /// Sends a POST request with the exception data to the web service
        /// at http://xltoolbox.sourceforge.net/receive.php
        /// </summary>
        public void Send()
        {
            using (WebClient client = new WebClient())
            {
                NameValueCollection v = new NameValueCollection(10);
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
                v["toolbox_version"] = ToolboxVersion;
                v["excel_version"] = ExcelVersion;
                v["excel_bitness"] = ExcelBitness;
                v["operating_system"] = OS;
                v["os_bitness"] = OSBitness;
                v["clr_version"] = CLR;
                v["freeimage_version"] = FreeImageVersion;
                client.UploadValuesCompleted += client_UploadValuesCompleted;
                client.UploadValuesAsync(new Uri(_postUrl), v);
            }
        }

        void client_UploadValuesCompleted(object sender, UploadValuesCompletedEventArgs e)
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
    }
}

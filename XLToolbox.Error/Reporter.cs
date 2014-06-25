using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        public string User { get; set; }
        public string Email { get; set; }
        public bool CcUser { get; set; }

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

        public string ReportID { get; private set; }

        /// <summary>
        /// Instantiates the class and sets the report ID to the hexadecimal
        /// representation of the current ticks (time elapsed since 1 AD).
        /// </summary>
        public Reporter(Application a, Exception e)
        {
            /* To produce a 'unique' error ID, we take the system time in ticks
             * elapsed since 1/1/2000, bit-shift it by 20 bits (equivalent to
             * dividing by roughly 1 million to get 1/10 of a second), then
             * converting it to a hexadecimal string represenation.
             */
            long baseDate = (new DateTime(2000, 1, 1)).Ticks >> 20;
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

    }
}

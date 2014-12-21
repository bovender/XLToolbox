using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bovender.Mvvm.Messaging
{
    public class FileNameMessageContent : StringMessageContent
    {
        #region Public properties

        public string Filter { get; set; }

        #endregion

        #region Public methods

        /// <summary>
        /// Add a file filter to the filter string.
        /// </summary>
        /// <param name="Description">Human-readable description</param>
        /// <param name="extensionWildcard">File extension including
        /// wildcard (e.g., "*.xlsx").</param>
        public void AddFilter(string Description, string extensionWildcard)
        {
            Filter += Description + "|" + extensionWildcard.ToUpper() + "|";
        }

        #endregion

        #region Constructor

        public FileNameMessageContent() 
            : base()
        {
            Filter = "";
        }

        public FileNameMessageContent(string filter)
            : base()
        {
            Filter = filter;
        }

        public FileNameMessageContent(string initialValue, string filter)
            : base(initialValue)
        {
            Filter = filter;
        }

        #endregion
    }
}

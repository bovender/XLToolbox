using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Error
{
    public class UploadFailedEventArgs : EventArgs
    {
        public Exception Error { get; private set; }

        public UploadFailedEventArgs(Exception error)
        {
            Error = error;
        }
    }
}

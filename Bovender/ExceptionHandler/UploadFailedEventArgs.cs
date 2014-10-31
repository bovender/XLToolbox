using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bovender.ExceptionHandler
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

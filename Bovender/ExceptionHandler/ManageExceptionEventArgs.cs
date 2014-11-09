using System;

namespace Bovender.ExceptionHandler
{
    public class ManageExceptionEventArgs : EventArgs
    {
        #region Public properties

        public Exception Exception { get; set; }
        public bool IsHandled { get; set; }

        #endregion

        #region Constructors

        public ManageExceptionEventArgs() { }

        public ManageExceptionEventArgs(Exception e)
            : this()
        {
            Exception = e;
        }

        #endregion
    }
}

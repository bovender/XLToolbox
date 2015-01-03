using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Export
{
    public class ExportProgressChangedEventArgs : EventArgs
    {
        #region Public properties

        public double PercentCompleted { get; set; }

        #endregion

        #region Constructor

        public ExportProgressChangedEventArgs() : base() { }

        public ExportProgressChangedEventArgs(double percentCompleted)
        {
            PercentCompleted = percentCompleted;
        }

        #endregion
    }
}

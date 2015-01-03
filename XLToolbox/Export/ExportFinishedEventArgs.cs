using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Export
{
    public class ExportFinishedEventArgs : EventArgs
    {
        #region Public properties

        public int FilesCreated { get; private set; }
        public bool WasCancelled { get; private set; }

        #endregion

        #region Constructor

        public ExportFinishedEventArgs(int filesCreated, bool wasCancelled)
        {
            FilesCreated = filesCreated;
            WasCancelled = WasCancelled;
        }

        #endregion
    }
}

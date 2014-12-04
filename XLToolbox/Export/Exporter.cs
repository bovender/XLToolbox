using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Bovender.Unmanaged;
using FreeImageAPI;

namespace XLToolbox.Export
{
    /// <summary>
    /// Provides methods to export the current selection from Excel.
    /// </summary>
    public class Exporter : IDisposable
    {
        #region Public methods

        /// <summary>
        /// Exports the current selection from Excel to a graphics file
        /// using the parameters defined in <see cref="exportSettings"/>
        /// </summary>
        /// <param name="exportSettings">Parameters for the graphic export.</param>
        /// <param name="fileName">Target file name.</param>
        public void ExportSelection(Settings exportSettings, string fileName)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Constructor and disposing

        public Exporter()
        {
            _freeimageDllHandle = DllManager.LoadLibrary("freeimage.dll");
        }

        ~Exporter()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool calledFromDispose)
        {
            DllManager.FreeLibrary(_freeimageDllHandle);
        }

        #endregion

        #region Private methods


        #endregion

        #region Private fields

        IntPtr _freeimageDllHandle;

        #endregion
    }
}

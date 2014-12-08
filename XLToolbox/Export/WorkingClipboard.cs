using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using Bovender.Unmanaged;

namespace XLToolbox.Export
{
    /// <summary>
    /// Replacement for the Sysem.Windows.Clipboard class that provides
    /// Metafile-related functions that actually work. Uses Windows P/Invoke!
    /// </summary>
    class WorkingClipboard : IDisposable
    {
        #region Public methods

        public Metafile GetMetafile()
        {
            _emfHandle = Pinvoke.GetClipboardData(Pinvoke.CF_ENHMETAFILE);
            return new Metafile(_emfHandle, true);
        }

        #endregion

        #region Constructor, disposal

        public WorkingClipboard()
        {
            Pinvoke.OpenClipboard((IntPtr)Excel.Instance.ExcelInstance.Application.Hwnd);
        }

        ~WorkingClipboard()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        void Dispose(bool calledFromDispose)
        {
            if (!_disposed)
            {
                _disposed = true;
                Pinvoke.CloseClipboard();
                if (_emfHandle != IntPtr.Zero)
                {
                    Pinvoke.DeleteEnhMetaFile(_emfHandle);
                }
            }
        }

        #endregion

        #region Private fields

        IntPtr _emfHandle;
        bool _disposed;

        #endregion
    }
}

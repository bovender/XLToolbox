/* WorkingClipboard.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2016 Daniel Kraus
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using Bovender.Extensions;
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
            try
            {
                Pinvoke.OpenClipboard((IntPtr)Excel.ViewModels.Instance.Default.Application.Hwnd);
            }
            catch (System.ComponentModel.Win32Exception e)
            {
                throw new WorkingClipboardException("Unable to obtain access to Windows clipboard", e);
            }
        }

        ~WorkingClipboard()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        void Dispose(bool calledFromDispose)
        {
            if (!_disposed)
            {
                _disposed = true;
                // Free unmanaged resources (Pinvoke class has static methods)
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

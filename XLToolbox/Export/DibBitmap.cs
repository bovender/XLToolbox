using FreeImageAPI;
/* DibBitmap.cs
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
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Bovender.Extensions;
using System.Text;

namespace XLToolbox.Export
{
    /// <summary>
    /// Provides access to a System.Drawing.Bitmap that is created
    /// from a DIB.
    /// </summary>
    /// <remarks>
    /// Inspired by:
    /// <list type="unordered">
    /// <item>http://hecgeek.blogspot.de/2007/04/converting-from-dib-to.html</item>
    /// <item>http://stackoverflow.com/questions/1054009/how-can-i-pass-memorystream-data-to-unmanaged-c-dll-using-p-invoke</item>
    /// <item>http://snipplr.com/view/36593.44712/</item>
    /// </list>
    /// </remarks>
    public class DibBitmap : IDisposable
    {
        #region Public properties

        public Bitmap Bitmap
        {
            get
            {
                if (_bitmap == null)
                {
                    _bitmap = new Bitmap(
                        Width, Height,
                        Width * 4,
                        // (Width * 24 + (32-1)) / 32 * 4,
                        PixelFormat.Format32bppRgb,
                        Scan0);
                }
                return _bitmap;
            }
        }

        public int Width
        {
            get
            {
                return DibHeader.biWidth;
            }
        }

        public int Height
        {
            get
            {
                return DibHeader.biHeight;
            }
        }

        #endregion

        #region Constructor

        public DibBitmap(Stream dibStream)
        {
            if (dibStream == null)
            {
                throw new ArgumentNullException("dibStream", "Stream with DIB data is required");
            }
            _stream = dibStream;
        }

        #endregion

        #region Dispose

        ~DibBitmap()
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
            if (!_disposed)
            {
                _disposed = true;
                if (calledFromDispose)
                {
                    if (_handle != null && _handle.IsAllocated)
                    {
                        _handle.Free();
                    }
                    if (_bitmap != null)
                    {
                        _bitmap.Dispose();
                    }
                    _stream.Dispose();
                }
            }
        }

        #endregion

        #region Private properties

        private byte[] DibBytes
        {
            get
            {
                if (_bytes == null)
                {
                    _bytes = new byte[_stream.Length];
                    _stream.Read(_bytes, 0, Convert.ToInt32(_stream.Length));
                }
                return _bytes;
            }
        }

        private System.Runtime.InteropServices.GCHandle DibHandle
        {
            get
            {
                if (_handle == null || !_handle.IsAllocated)
                {
                    _handle = System.Runtime.InteropServices.GCHandle.Alloc(
                        DibBytes,
                        System.Runtime.InteropServices.GCHandleType.Pinned);
                }
                return _handle;
            }
        }

        private IntPtr Scan0
        {
            get
            {
                if (_scan0 == IntPtr.Zero)
                {
                    IntPtr handlePtr = DibHandle.AddrOfPinnedObject();
                    if (Environment.Is64BitProcess)
                    {
                        _scan0 = new IntPtr(handlePtr.ToInt64() + 40);
                    }
                    else
                    {
                        _scan0 = new IntPtr(handlePtr.ToInt32() + 40);
                    }
                }
                return _scan0;
            }
        }

        private BITMAPINFOHEADER DibHeader
        {
            get
            {
                if (!_headerInitialized)
                {
                    _headerInitialized = true;
                    _dibHeader = (BITMAPINFOHEADER)System.Runtime.InteropServices.Marshal.PtrToStructure(
                        DibHandle.AddrOfPinnedObject(), _dibHeader.GetType());
                }
                return _dibHeader;
            }
        }

        #endregion

        #region Private fields

        private bool _disposed;
        private Bitmap _bitmap;
        private Stream _stream;
        private byte[] _bytes;
        private IntPtr _scan0;
        private BITMAPINFOHEADER _dibHeader;
        private bool _headerInitialized;
        private System.Runtime.InteropServices.GCHandle _handle;

        #endregion
    }
}

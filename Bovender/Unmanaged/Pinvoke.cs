using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Bovender.Unmanaged
{
    /// <summary>
    /// Wrappers for Win32 API calls.
    /// </summary>
    public static class Pinvoke
    {
        #region Public methods

        public static void OpenClipboard(IntPtr hWndNewOwner)
        {
            if (!Win32_OpenClipboard(hWndNewOwner))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }

        public static void CloseClipboard()
        {
            if (!Win32_CloseClipboard())
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }

        public static IntPtr GetClipboardData(uint uFormat)
        {
             IntPtr result = Win32_GetClipboardData(uFormat);
             if (result == IntPtr.Zero)
             {
                 throw new Win32Exception(Marshal.GetLastWin32Error());
             }
             return result;
        }

        public static IntPtr CopyEnhMetaFile(IntPtr hemfSrc, string lpszFile)
        {
            IntPtr result = Win32_CopyEnhMetaFile(hemfSrc, lpszFile);
            if (result == IntPtr.Zero)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            return result;
        }

        public static void DeleteEnhMetaFile(IntPtr hemf)
        {
            if (!Win32_DeleteEnhMetaFile(hemf))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }

        #endregion

        #region Win32 API constants

        public const uint CF_ENHMETAFILE = 14;

        #endregion

        #region Win32 DLL imports

        [DllImport("user32.dll", EntryPoint = "OpenClipboard", SetLastError = true)]
        static extern bool Win32_OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", EntryPoint = "CloseClipboard", SetLastError = true)]
        static extern bool Win32_CloseClipboard();

        [DllImport("user32.dll", EntryPoint = "GetClipboardData", SetLastError = true)]
        static extern IntPtr Win32_GetClipboardData(uint uFormat);

        [DllImport("gdi32.dll", EntryPoint = "CopyEnhMetaFile", SetLastError = true)]
        static extern IntPtr Win32_CopyEnhMetaFile(IntPtr hemfSrc, string lpszFile);

        [DllImport("gdi32.dll", EntryPoint = "DeleteEnhMetaFile", SetLastError = true)]
        static extern bool Win32_DeleteEnhMetaFile(IntPtr hemf);

        #endregion
    }
}

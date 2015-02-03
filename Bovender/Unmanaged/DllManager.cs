/* DllManager.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
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
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace Bovender.Unmanaged
{
    /// <summary>
    /// Manages unmanaged DLLs. Unloads any loaded DLLs upon disposal.
    /// </summary>
    public class DllManager : Object, IDisposable
    {
        #region constants
        private const string LIBDIR = "lib";
        #endregion

        #region WinAPI

        [DllImport("kernel32.dll", EntryPoint="LoadLibrary", SetLastError=true)]
        static extern IntPtr LoadLibrary(string dllToLoad);

        [DllImport("kernel32.dll", EntryPoint="GetProcAddress", SetLastError=true)]
        static extern IntPtr GetProcAddress(IntPtr hModule, string procedureName);

        [DllImport("kernel32.dll", EntryPoint = "FreeLibrary", SetLastError = true)]
        static extern bool FreeLibrary(IntPtr hModule);

        #endregion
        
        #region Private members
        /// <summary>
        /// Holds the currently loaded DLL names and their handles.
        /// </summary>
        private Dictionary<string, IntPtr> _loadedDlls = new Dictionary<string, IntPtr>();

        private bool _disposed;
        #endregion

        #region Loading and unloading DLLs

        /// <summary>
        /// Loads the given DLL from the appropriate subdirectory, depending on the current
        /// bitness.
        /// </summary>
        /// <remarks>
        /// The DLL is expected to reside in a subdirectory of the entry point assembly's
        /// directory, in "bin/lib/$(Platform)", where $(Platform) can be "Win32" or "x64",
        /// for example.
        /// </remarks>
        /// <exception cref="DllNotFoundException">if the file is not found in the path.</exception>
        /// <exception cref="DllLoadingFailedException">if the file is not found in the path.</exception>
        /// <param name="dllName">Name of the DLL to load (without path).</param>
        public void LoadDll(string dllName)
        {
            // Check if the DLL exists
            string dllPath = CompletePath(dllName);
            CheckFilePresent(dllPath);

            // Attempt to load the DLL
            IntPtr handle = LoadLibrary(dllPath);
            if (handle == IntPtr.Zero)
            {
                // Strip the leading directories from the path info (they may contain
                // sensitive information about where exactly a user has installed files).
                string[] dirs = Path.GetDirectoryName(dllPath).Split(Path.DirectorySeparatorChar);
                string gracefulPath = dllName;
                int n = dirs.Length;
                if (n > 0) gracefulPath = Path.Combine(dirs[n - 1], gracefulPath);
                if (n > 1) gracefulPath = Path.Combine(dirs[n - 2], gracefulPath);
                if (n > 2) gracefulPath = Path.Combine("...", gracefulPath);
                Win32Exception inner = new Win32Exception(Marshal.GetLastWin32Error());
                throw new DllLoadingFailedException(
                    String.Format(
                        "LoadLibrary failed with code {0} on {1}",
                        Marshal.GetLastWin32Error(), gracefulPath
                    ),
                    inner
                );
            }

            // Register the DLL and its handle in the internal database
            _loadedDlls.Add(dllName, handle);
        }

        /// <summary>
        /// Loads the given DLL from the appropriate subdirectory if its Sha1 hash
        /// matches the provided hash.
        /// </summary>
        /// <param name="dllName">Name of the DLL to load (without path).</param>
        /// <param name="expectedSha1Hash">Expected Sha1 hash of the DLL.</param>
        /// <exception cref="DllNotFoundException">if the file is not found in the path.</exception>
        /// <exception cref="DllLoadingFailedException">if the file is not found in the path.</exception>
        /// <exception cref="DllSha1MismatchException">if the file's Sha1 is unexpected.</exception>
        // TODO: Use two expected hashes, one for Win32, one for x64
        public void LoadDll(string dllName, string expectedSha1Hash)
        {
            string dllPath = CompletePath(dllName);
            CheckFilePresent(dllPath);
            string actualSha1Hash = FileHelpers.Sha1Hash(dllPath);
            if (actualSha1Hash != expectedSha1Hash)
            {
                throw new DllSha1MismatchException(String.Format(
                    "Expected {0} but got {1} on {2}", expectedSha1Hash, actualSha1Hash, dllPath));
            };
            LoadDll(dllName);
        }

        /// <summary>
        /// Unloads a previously loaded DLL. Does nothing if the DLL was not loaded.
        /// </summary>
        /// <param name="dllName">Name of the DLL to unload.</param>
        public void UnloadDll(string dllName)
        {
            IntPtr handle;
            if (_loadedDlls.TryGetValue(dllName, out handle))
            {
                FreeLibrary(handle);
            }
        }

        #endregion

        #region Destructing and disposing

        ~DllManager()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                Dispose(true);
                GC.SuppressFinalize(this);
                _disposed = true;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                foreach (IntPtr handle in _loadedDlls.Values)
                {
                    FreeLibrary(handle);
                }
            }
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Checks if the given file exists and throws a custom exception if not.
        /// </summary>
        /// <remarks>
        /// Since an exception will be thrown by the CLR anyway, this method may seem
        /// redundant. However, throwing a custom exception might help to identify the
        /// problem a user has if only the name of the exception is reported and the
        /// call trace is unknown.
        /// </remarks>
        /// <exception cref="DllNotFoundException">if file was not found</exception>
        /// <param name="file">File whose presence to check.</param>
        private void CheckFilePresent(string file)
        {
            if (!System.IO.File.Exists(file))
            {
                throw new DllNotFoundException(String.Format(
                    "Not found: {0}", file));
            };
        }

        /// <summary>
        /// Constructs and returns the complete path for a given DLL.
        /// </summary>
        /// <remarks>
        /// By convention, the path that the DLL is expected to reside in is a subdirectory
        /// of the entry point assembly's base directory, in "bin/lib/$(Platform)", where
        /// $(Platform) can be "Win32" or "x64", for example.
        /// </remarks>
        /// <param name="fileName">Name of the DLL (with or without extension).</param>
        /// <returns>Path to the DLL subdirectory (platform-dependent).</returns>
        private string CompletePath(string fileName)
        {
            if (!Path.HasExtension(fileName))
            {
                fileName += ".dll";
            };
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "lib",
                Environment.Is64BitProcess ? "x64" : "Win32",
                fileName);
        }

        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows;
using Bovender.Unmanaged;
using FreeImageAPI;
using XLToolbox.Excel.Instance;
using XLToolbox.Excel.ViewModels;
using System.Runtime.InteropServices;

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
            // Copy current selection to clipboard
            SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
            svm.CopyToClipboard();
            WorkingClipboard clipboard = new WorkingClipboard();
            Metafile emf = clipboard.GetMetafile();
            string tempFile = System.IO.Path.GetTempFileName();
            // Analyze size of selection
            // Determine bpp
            int px = 4000;
            int py = 3000;
            Bitmap b = new Bitmap(px, py);
            Graphics g = Graphics.FromImage(b);
            g.FillRectangle(Brushes.White, 0, 0, px, py);
            g.DrawImage(emf, 0, 0, px, py);
            // b.MakeTransparent(Color.White);
            FreeImageBitmap fib = new FreeImageBitmap(b);
            // Convert color space
            // Attach color profile
            fib.SetResolution(exportSettings.Dpi, exportSettings.Dpi);
            fib.Save(fileName);
        }

        #endregion

        #region Constructor and disposing

        public Exporter()
        {
            _dllManager = new DllManager();
            _dllManager.LoadDll("freeimage.dll");
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
            if (calledFromDispose && !_disposed)
            {
                _dllManager.UnloadDll("freeimage.dll");
                _disposed = true;
            }
        }

        #endregion

        #region Protected properties
        #endregion

        #region Private methods

        /// <summary>
        /// Analyzes the dimensions of the current selection in Excel
        /// </summary>
        private void AnalyzeSelection()
        {
            object selection = ExcelInstance.Application.Selection;

        }

        #endregion

        #region Private fields

        DllManager _dllManager;
        bool _disposed;

        #endregion

    }
}

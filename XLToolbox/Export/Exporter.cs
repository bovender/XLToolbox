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
        #region Public properties

        /// <summary>
        /// The export preset (file type, resolution, color space)
        /// to use for the graphic exports.
        /// </summary>
        public Preset Preset { get; set; }

        #endregion

        #region Public methods

        /// <summary>
        /// Exports the current selection from Excel to a graphics file
        /// using the parameters defined in <see cref="exportSettings"/>
        /// </summary>
        /// <param name="exportSettings">Parameters for the graphic export.</param>
        /// <param name="fileName">Target file name.</param>
        public void ExportSelection(SingleExportSettings settings)
        {
            if (Preset == null)
            {
                throw new InvalidOperationException("No Preset given.");
            }
            if (settings == null)
            {
                throw new ArgumentNullException("settings",
                    "Must have SingleExportSettings object for the export.");
            }

            DoExportSelection(settings.Width, settings.Height, settings.FileName);
        }

        #endregion

        #region Constructor and disposing

        public Exporter()
        {
            _dllManager = new DllManager();
            _dllManager.LoadDll("freeimage.dll");
        }

        public Exporter(Preset preset)
            : this()
        {
            Preset = preset;
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
        /// Performs the actual export for a given selection. This method is
        /// called by the <see cref="ExportSelection"/> method and during
        /// a batch export process.
        /// </summary>
        /// <param name="width">Width of the output graphic.</param>
        /// <param name="height">Height of the output graphic.</param>
        /// <param name="fileName">Destination filename (must contain placeholders).</param>
        private void DoExportSelection(double width, double height, string fileName)
        {
            // Copy current selection to clipboard
            SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
            svm.CopyToClipboard();

            // Get a metafile view of the clipboard content
            WorkingClipboard clipboard = new WorkingClipboard();
            Metafile emf = clipboard.GetMetafile();

            // Calculate the number of pixels needed for the requested
            // output size and resolution; size is given in points (1/72 in),
            // resolution is given in dpi.
            int px = (int)Math.Round(width / 72 * Preset.Dpi);
            int py = (int)Math.Round(height / 72 * Preset.Dpi);

            // Draw the image on a GDI+ bitmap
            Bitmap b = new Bitmap(px, py);
            Graphics g = Graphics.FromImage(b);
            g.FillRectangle(Brushes.White, 0, 0, px, py);
            g.DrawImage(emf, 0, 0, px, py);
            b.MakeTransparent(Color.White);

            // Create a FreeImage bitmap from the GDI+ bitmap
            FreeImageBitmap fib = new FreeImageBitmap(b);
            // TODO: Convert color space
            // TODO: Attach color profile
            fib.SetResolution(Preset.Dpi, Preset.Dpi);
            fib.Save(fileName);
        }

        #endregion

        #region Private fields

        DllManager _dllManager;
        bool _disposed;

        #endregion
    }
}

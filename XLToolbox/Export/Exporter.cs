﻿/* Exporter.cs
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
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Office.Interop.Excel;
using Bovender.Unmanaged;
using FreeImageAPI;
using FreeImageAPI.Metadata;
using XLToolbox.Excel.Instance;
using XLToolbox.Excel.ViewModels;
using System.Collections.Generic;
using XLToolbox.Export.Models;
using System.Threading.Tasks;

namespace XLToolbox.Export
{
    /// <summary>
    /// Provides methods to export the current selection from Excel.
    /// </summary>
    public class Exporter : IDisposable
    {
        #region Events

        /// <summary>
        /// Signals current export progress during batch export operations.
        /// </summary>
        public event EventHandler<ExportProgressChangedEventArgs> BatchExportProgressChanged;

        /// <summary>
        /// Signals that an asynchronous export progress is finished. Subscribers
        /// can inspect the event args to find out if the process has been cancelled
        /// prematurely.
        /// </summary>
        public event EventHandler<ExportFinishedEventArgs> BatchExportFinished;

        #endregion

        #region Public methods

        /// <summary>
        /// Performs a quick export using a given Preset, but
        /// without altering the size of the current selection:
        /// The dimension properties of the SingleExportSettings
        /// object that defines the operation are ignored.
        /// </summary>
        /// <param name="preset"></param>
        public void ExportSelectionQuick(SingleExportSettings settings)
        {
            ExportSelection(settings.Preset, settings.FileName);
        }

        /// <summary>
        /// Exports the current selection from Excel to a graphics file
        /// using the parameters defined in <see cref="exportSettings"/>
        /// </summary>
        /// <param name="exportSettings">Parameters for the graphic export.</param>
        /// <param name="fileName">Target file name.</param>
        public void ExportSelection(SingleExportSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings",
                    "Must have SingleExportSettings object for the export.");
            }
            double w = settings.Unit.ConvertTo(settings.Width, Unit.Point);
            double h = settings.Unit.ConvertTo(settings.Height, Unit.Point);
            ExportSelection(settings.Preset, w, h, settings.FileName);
        }

        /// <summary>
        /// Starts a batch export of charts and/or drawing objects
        /// (shapes) asynchronously. Callers should subscribe to the
        /// <see cref="ExportProcessChanged"/> and <see cref="ExportFinished"/>
        /// events to learn about status changes. The operation can
        /// be cancelled by caling the <see cref="CancelBatchExport"/>
        /// method.
        /// </summary>
        /// <param name="settings">Settings describing the desired operation.</param>
        public void ExportBatchAsync(BatchExportSettings settings)
        {
            if (_batchRunning)
            {
                throw new InvalidOperationException(
                    "Cannot start batch export while an operation is in progress.");
            }

            Task t = new Task(() =>
            {
                _batchRunning = true;
                ExcelInstance.DisableScreenUpdating();
                switch (_batchSettings.Scope)
                {
                    case BatchExportScope.ActiveSheet:
                        _numTotal = CountInSheet(ExcelInstance.Application.ActiveSheet);
                        ExportSheet(ExcelInstance.Application.ActiveSheet);
                        break;
                    case BatchExportScope.ActiveWorkbook:
                        _numTotal = CountInWorkbook(ExcelInstance.Application.ActiveWorkbook);
                        ExportWorkbook(ExcelInstance.Application.ActiveWorkbook);
                        break;
                    case BatchExportScope.OpenWorkbooks:
                        _numTotal = CountInAllWorkbooks();
                        ExportAllWorkbooks();
                        break;
                    default:
                        throw new NotImplementedException(String.Format(
                            "Batch export not implemented for {0}",
                            settings.Scope));
                }
                ExcelInstance.EnableScreenUpdating();
                OnExportFinished();
                _batchRunning = false;
            });

            _batchSettings = settings;
            _cancelled = false;
            _batchFileName = new ExportFileName(settings.Path, settings.FileName,
                settings.Preset.FileType);
            t.Start();
        }

        /// <summary>
        /// Cancels a running batch export.
        /// </summary>
        public void CancelBatchExport()
        {
            if (_batchRunning) _cancelled = true;
        }

        #endregion

        #region Constructor and disposing

        public Exporter()
        {
            _dllManager = new DllManager();
            _dllManager.LoadDll("freeimage.dll");
            _fileTypeToFreeImage = new Dictionary<FileType, FREE_IMAGE_FORMAT>()
            {
                { FileType.Png, FREE_IMAGE_FORMAT.FIF_PNG },
                { FileType.Tiff, FREE_IMAGE_FORMAT.FIF_TIFF }
            };
        }

        /*
        public Exporter(Preset preset)
            : this()
        {
            Preset = preset;
        }
         * */

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

        #region Private export methods

        /// <summary>
        /// Performs the actual graphic export with the dimensions of
        /// the current selection.
        /// </summary>
        /// <param name="preset">Export preset to use.</param>
        /// <param name="fileName">File name of target file.</param>
        private void ExportSelection(Preset preset, string fileName)
        {
            SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
            if (svm.Selection == null)
            {
                throw new InvalidOperationException("Nothing selected in Excel.");
            }
            ExportSelection(preset, svm.Bounds.Width, svm.Bounds.Height, fileName);
        }

        /// <summary>
        /// Performs the actual export for a given selection. This method is
        /// called by the <see cref="ExportSelection"/> method and during
        /// a batch export process.
        /// </summary>
        /// <param name="widthInPoints">Width of the output graphic.</param>
        /// <param name="heightInPoints">Height of the output graphic.</param>
        /// <param name="fileName">Destination filename (must contain placeholders).</param>
        private void ExportSelection(Preset preset, double widthInPoints, double heightInPoints,
            string fileName)
        {
            // Copy current selection to clipboard
            SelectionViewModel svm = new SelectionViewModel(ExcelInstance.Application);
            svm.CopyToClipboard();

            // Get a metafile view of the clipboard content
            // Must not dispose the WorkingClipboard instance before the metafile
            // has been drawn on the bitmap canvas! Otherwise the metafile will not draw.
            Metafile emf;
            using (WorkingClipboard clipboard = new WorkingClipboard())
            {
                emf = clipboard.GetMetafile();
                switch (preset.FileType)
                {
                    case FileType.Emf:
                        ExportEmf(emf, fileName);
                        break;
                    case FileType.Png:
                    case FileType.Tiff:
                        ExportViaFreeImage(emf, preset, widthInPoints, heightInPoints, fileName);
                        break;
                    default:
                        throw new NotImplementedException(String.Format(
                            "No export implementation for {0}.", preset.FileType));
                }
            }
        }

        private void ExportViaFreeImage(Metafile metafile,
            Preset preset, double width, double height, string fileName)
        {
            // Calculate the number of pixels needed for the requested
            // output size and resolution; size is given in points (1/72 in),
            // resolution is given in dpi.
            int px = (int)Math.Round(width / 72 * preset.Dpi);
            int py = (int)Math.Round(height / 72 * preset.Dpi);

            // Create a canvas (GDI+ bitmap) and associate it with a
            // Graphics object.
            Bitmap b = new Bitmap(px, py);
            Graphics g = Graphics.FromImage(b);

            // Get a brush to paint the canvas
            Brush brush;
            if (preset.Transparency == Transparency.TransparentCanvas)
            {
                brush = Brushes.Transparent;
            }
            else
            {
                brush = Brushes.White;
            }
            g.FillRectangle(brush, 0, 0, px, py);

            // Draw the image on the canvas
            g.DrawImage(metafile, 0, 0, px, py);

            // Make the white colors transparent if required
            if (preset.Transparency == Transparency.TransparentWhite)
            {
                b.MakeTransparent(Color.White);
            }

            // Create a FreeImage bitmap from the GDI+ bitmap
            FreeImageBitmap fib = new FreeImageBitmap(b);

            fib.ConvertColorDepth(preset.ColorSpace.ToFreeImageColorDepth());
            if (preset.ColorSpace == ColorSpace.Monochrome)
            {
                fib.Palette.SetValue(new RGBQUAD(Color.Black), 0);
                fib.Palette.SetValue(new RGBQUAD(Color.White), 1);
            }

            // TODO: Attach color profile
            fib.SetResolution(preset.Dpi, preset.Dpi);

            fib.Comment = Versioning.SemanticVersion.BrandName;
            fib.Save(
                fileName,
                preset.FileType.ToFreeImageFormat(),
                GetSaveFlags(preset)
            );
        }

        private void ExportEmf(Metafile metafile, string fileName)
        {
            metafile.Save(fileName);
        }

        private void ExportAllWorkbooks()
        {
            foreach (Workbook wb in ExcelInstance.Application.Workbooks)
            {
                ExportWorkbook(wb);
                if (_cancelled) break;
            }
        }

        private void ExportWorkbook(Workbook workbook)
        {
            ((_Workbook)workbook).Activate();
            foreach (dynamic ws in workbook.Sheets)
            {
                ExportSheet(ws);
                if (_cancelled) break;
            }
        }

        private void ExportSheet(dynamic sheet)
        {
            sheet.Activate();
            switch (_batchSettings.Layout)
            {
                case BatchExportLayout.SheetLayout:
                    ExportSheetLayout(sheet);
                    break;
                case BatchExportLayout.SingleItems:
                    ExportSheetItems(sheet);
                    break;
                default:
                    throw new NotImplementedException(
                        String.Format("Export of {0} not implemented.", _batchSettings.Layout)
                        );
            }
        }

        private void ExportSheetLayout(dynamic sheet)
        {
            SheetViewModel svm = new SheetViewModel(sheet);
            switch (_batchSettings.Objects)
            {
                case BatchExportObjects.Charts:
                    svm.SelectCharts();
                    break;
                case BatchExportObjects.ChartsAndShapes:
                    svm.SelectShapes();
                    break;
                default:
                    throw new NotImplementedException(_batchSettings.Objects.ToString());
            }
            ExportSelection(
                _batchSettings.Preset,
                _batchFileName.GenerateNext(sheet)
            );
            OnExportProgressChanged();
        }

        private void ExportSheetItems(dynamic sheet)
        {
            SheetViewModel svm = new SheetViewModel(sheet);
            if (svm.IsChart)
            {
                svm.SelectCharts();
                ExportSelection(
                    _batchSettings.Preset,
                    _batchFileName.GenerateNext(sheet)
                );
            }
            else
            {
                switch (_batchSettings.Objects)
                {
                    case BatchExportObjects.Charts:
                        ExportSheetChartItems(svm.Worksheet);
                        break;
                    case BatchExportObjects.ChartsAndShapes:
                        ExportSheetAllItems(svm.Worksheet);
                        break;
                    default:
                        throw new NotImplementedException(
                            "Single-item export not implemented for " + _batchSettings.Objects.ToString());
                }
            }
            OnExportProgressChanged();
        }

        private void ExportSheetChartItems(Worksheet worksheet)
        {
            // Must use an index-based for loop here.
            // A foreach loop caused lots of 0x800a03ec errors from Excel
            // (for whatever reason).
            ChartObjects cos = worksheet.ChartObjects();
            for (int i = 1; i <= cos.Count; i++)
            {
                cos.Item(i).Select();
                ExportSelection(_batchSettings.Preset, _batchFileName.GenerateNext(worksheet));
                OnExportProgressChanged();
                if (_cancelled) break;
            }
        }

        private void ExportSheetAllItems(Worksheet worksheet)
        {
            foreach (Shape sh in worksheet.Shapes)
            {
                sh.Select(true);
                ExportSelection(_batchSettings.Preset, _batchFileName.GenerateNext(worksheet));
                OnExportProgressChanged();
                if (_cancelled) break;
            }
        }

        #endregion

        #region Private counting methods

        private int CountInAllWorkbooks()
        {
            int n = 0;
            foreach (Workbook wb in ExcelInstance.Application.Workbooks)
            {
                n += CountInWorkbook(wb);
            }
            return n;
        }

        private int CountInWorkbook(Workbook workbook)
        {
            int n = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                n += CountInSheet(ws);
            }
            return n;
        }

        private int CountInSheet(dynamic worksheet)
        {
            switch (_batchSettings.Layout)
            {
                case BatchExportLayout.SheetLayout:
                    return CountInSheetLayout(worksheet);
                case BatchExportLayout.SingleItems:
                    return CountInSheetItems(worksheet);
                default:
                    throw new NotImplementedException(
                        String.Format("Export of {0} not implemented.", _batchSettings.Layout)
                        );
            }
        }

        /// <summary>
        /// Returns 1 if the <paramref name="worksheet"/> contains at least
        /// one chart or drawing object, since all charts/drawing objects will
        /// be exported together into one file.
        /// </summary>
        /// <param name="worksheet">Worksheet to examine.</param>
        /// <returns>1 if sheet contains charts/drawings, 0 if not.</returns>
        private int CountInSheetLayout(dynamic worksheet)
        {
            SheetViewModel svm = new SheetViewModel(worksheet);
            switch (_batchSettings.Objects)
            {
                case BatchExportObjects.Charts:
                    return svm.CountCharts() > 0 ? 1 : 0;
                case BatchExportObjects.ChartsAndShapes:
                    return svm.CountShapes() > 0 ? 1 : 0;
                default:
                    throw new NotImplementedException(String.Format(
                        "Export of {0} not implemented.", _batchSettings.Objects));
            }
        }

        private int CountInSheetItems(dynamic worksheet)
        {
            SheetViewModel svm = new SheetViewModel(worksheet);
            switch (_batchSettings.Objects)
            {
                case BatchExportObjects.Charts:
                    return svm.CountCharts();
                case BatchExportObjects.ChartsAndShapes:
                    return svm.CountShapes();
                default:
                    throw new NotImplementedException(String.Format(
                        "Export of {0} not implemented.", _batchSettings.Objects));
            }
        }

        /*private FREE_IMAGE_FORMAT FileTypeToFreeImage(FileType fileType)
        {
            FREE_IMAGE_FORMAT fif;
            if (_fileTypeToFreeImage.TryGetValue(fileType, out fif))
            {
                return fif;
            }
            else
            {
                throw new NotImplementedException(
                    "No FREE_IMAGE_FORMAT match for " + fileType.ToString());
            }
        }
        */
        #endregion

        #region Private helper methods

        private void OnExportProgressChanged()
        {
            if (BatchExportProgressChanged != null)
            {
                BatchExportProgressChanged(
                    this,
                    new ExportProgressChangedEventArgs(
                        Convert.ToDouble(_batchFileName.Counter) / _numTotal)
                );
            }
        }

        private void OnExportFinished()
        {
            if (BatchExportFinished != null)
            {
                BatchExportFinished(this, new ExportFinishedEventArgs(
                    _batchFileName.Counter, _cancelled));
            }
        }

        private FREE_IMAGE_SAVE_FLAGS GetSaveFlags(Preset preset)
        {
            switch (preset.FileType)
            {
                case FileType.Png:
                    return FREE_IMAGE_SAVE_FLAGS.PNG_Z_BEST_COMPRESSION |
                           FREE_IMAGE_SAVE_FLAGS.PNG_INTERLACED;
                case FileType.Tiff:
                    switch (preset.ColorSpace)
                    {
                        case ColorSpace.Monochrome:
                            return FREE_IMAGE_SAVE_FLAGS.TIFF_CCITTFAX4;
                        default:
                            return FREE_IMAGE_SAVE_FLAGS.TIFF_LZW;
                    }
                default:
                    return FREE_IMAGE_SAVE_FLAGS.DEFAULT;
            }
        }

        /*
        /// <summary>
        /// Adds a file extension to the file name if missing.
        /// </summary>
        /// <param name="fileName">File name, possibly without extension.</param>
        /// <returns>File name with extension.</returns>
        private string SanitizeFileName(Preset preset, string fileName)
        {
            string extension = preset.FileType.ToFileNameExtension();
            if (!fileName.ToUpper().EndsWith(extension.ToUpper()))
            {
                fileName += extension;
            }
            return fileName;
        }
         */

        #endregion

        #region Private fields

        DllManager _dllManager;
        bool _disposed;
        BatchExportSettings _batchSettings;
        ExportFileName _batchFileName;
        bool _batchRunning;
        bool _cancelled;
        int _numTotal;
        Dictionary<FileType, FREE_IMAGE_FORMAT> _fileTypeToFreeImage;

        #endregion

        #region Private constants
        #endregion
    }
}
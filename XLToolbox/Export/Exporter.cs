﻿/* Exporter.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using FreeImageAPI;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Export.Models;
using Bovender.Extensions;
using Bovender.Unmanaged;

namespace XLToolbox.Export
{
    /// <summary>
    /// Provides methods to export the current selection from Excel.
    /// </summary>
    public class Exporter : Bovender.Mvvm.Models.ProcessModel, IExporter, IDisposable
    {
        #region Properties

        public string FileName { get; set; }

        /// <summary>
        /// Gets or sets whether to do a quick export or not.
        /// Quick export means that the selection is exported
        /// at the original size.
        /// </summary>
        public bool QuickExport { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public Preset Preset
        {
            get
            {
                if (_settings != null)
                {
                    Logger.Info("Preset_get: Giving preference to Settings object");
                    return _settings.Preset;
                }
                else
                {
                    Logger.Info("Preset_get: Setting private field");
                    return _preset;
                }
            }
            set
            {
                if (_settings != null)
                {
                    Logger.Info("Preset_set: Giving preference to Settings object");
                    _settings.Preset = value;
                }
                else
                {
                    Logger.Info("Preset_set: Setting private field");
                    _preset = value;
                }
            }
        }

        public int PercentCompleted
        {
            get
            {
                if (_tiledBitmap != null)
                {
                    return _tiledBitmap.PercentCompleted * 50 / 100 + _percentCompleted;
                }
                else
                {
                    return _percentCompleted;
                }
            }
            set
            {
                Logger.Info("percent: {0}", _percentCompleted);
                _percentCompleted = value;
            }
        }

        #endregion

        #region Public methods

        public override bool Execute()
        {
            if (Preset == null)
            {
                Logger.Fatal("Execute: No preset!");
                throw new InvalidOperationException("Execute: Cannot export because no Preset was given");
            }
            if (String.IsNullOrWhiteSpace(FileName) && _settings != null)
            {
                FileName = _settings.FileName;
            }
            if (String.IsNullOrWhiteSpace(FileName))
            {
                Logger.Fatal("Execute: No filename!");
                throw new InvalidOperationException("Execute: Cannot export because no file name was given");
            }
            bool result = false;
            double width;
            double height;
            if (QuickExport)
            {
                if (SelectionViewModel.Selection == null)
                {
                    Logger.Fatal("Execute: Quick export: No selection!");
                    throw new InvalidOperationException("Execute: Cannot export because nothing is selected in Excel");
                }
                width = SelectionViewModel.Bounds.Width;
                height = SelectionViewModel.Bounds.Height;
            }
            else
            {
                if (_settings == null)
                {
                    Logger.Fatal("Execute: No export settings!");
                    throw new InvalidOperationException("Cannot export because no export settings were given; want to perform quick export?");
                }
                width = _settings.Unit.ConvertTo(_settings.Width, Unit.Point);
                height = _settings.Unit.ConvertTo(_settings.Height, Unit.Point);
            }
            ExportWithDimensions(width, height);
            return result;
        }

        #endregion

        #region Constructors

        public Exporter(SingleExportSettings settings)
            : this()
        {
            _settings = settings;
            if (_settings != null)
	        {
                Preset = _settings.Preset;
	        }
        }

        public Exporter(SingleExportSettings settings, bool quickExport)
            : this(settings)
        {
            QuickExport = quickExport;
        }

        public Exporter(Preset preset)
            : this()
        {
            // Without SingleExportSettings, we can only perform a quick export
            QuickExport = true;
            Preset = preset;
        }

        protected Exporter()
            : base()
        {
            _dllManager = new DllManager();
            _dllManager.LoadDll("freeimage.dll");
            _fileTypeToFreeImage = new Dictionary<FileType, FREE_IMAGE_FORMAT>()
            {
                { FileType.Png, FREE_IMAGE_FORMAT.FIF_PNG },
                { FileType.Tiff, FREE_IMAGE_FORMAT.FIF_TIFF }
            };
        }

        #endregion

        #region Disposing

        ~Exporter()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected void Dispose(bool calledFromDispose)
        {
            if (!_disposed)
            {
                if (calledFromDispose)
                {
                    // Free managed resources
                    _dllManager.UnloadDll("freeimage.dll");
                    if (_tiledBitmap != null)
                    {
                        _tiledBitmap.Dispose();
                    }
                }
                _disposed = true;
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Performs the actual export for a given selection. This method is
        /// called by the <see cref="ExportSelection"/> method and during
        /// a batch export process.
        /// </summary>
        /// <param name="widthInPoints">Width of the output graphic.</param>
        /// <param name="heightInPoints">Height of the output graphic.</param>
        /// <param name="fileName">Destination filename (must contain placeholders).</param>
        private void ExportWithDimensions(double widthInPoints, double heightInPoints)
        {
            if (Preset == null)
            {
                Logger.Fatal("ExportWithDimensions: No export preset!");
                throw new InvalidOperationException("Cannot export without export preset");
            }
            Logger.Info("ExportWithDimensions: Preset: {0}", Preset);
            // Copy current selection to clipboard
            SelectionViewModel.CopyToClipboard();
                    
            // Get a metafile view of the clipboard content
            // Must not dispose the WorkingClipboard instance before the metafile
            // has been drawn on the bitmap canvas! Otherwise the metafile will not draw.
            Metafile emf;
            using (WorkingClipboard clipboard = new WorkingClipboard())
            {
                Logger.Info("ExportWithDimensions: Get metafile");
                emf = clipboard.GetMetafile();
                switch (Preset.FileType)
                {
                    case FileType.Emf:
                        ExportEmf(emf);
                        break;
                    case FileType.Png:
                    case FileType.Tiff:
                        ExportViaFreeImage(emf, widthInPoints, heightInPoints);
                        break;
                    default:
                        throw new NotImplementedException(String.Format(
                            "No export implementation for {0}.", Preset.FileType));
                }
            }
        }

        private void ExportViaFreeImage(Metafile metafile, double width, double height)
        {
            Logger.Info("ExportViaFreeImage: Preset: {0}", Preset);
            Logger.Info("ExportViaFreeImage: Width: {0}; height: {1}", width, height);
            // Calculate the number of pixels needed for the requested
            // output size and resolution; size is given in points (1/72 in),
            // resolution is given in dpi.
            int px = (int)Math.Round(width / 72 * Preset.Dpi);
            int py = (int)Math.Round(height / 72 * Preset.Dpi);
            Logger.Info("ExportViaFreeImage: Pixels: x: {0}; y: {1}", px, py);
            Cancelling += Exporter_Cancelling;
            PercentCompleted = 10;
            _tiledBitmap = new TiledBitmap(px, py);
            FreeImageBitmap fib = _tiledBitmap.CreateFreeImageBitmap(metafile, EffectiveTransparency());
            ConvertColor(fib);
            fib.SetResolution(Preset.Dpi, Preset.Dpi);
            fib.Comment = Versioning.SemanticVersion.Current.BrandName;
            PercentCompleted = 30;
            Logger.Info("ExportViaFreeImage: Saving {0} file", Preset.FileType);
            try
            {
                fib.Save(
                    FileName,
                    Preset.FileType.ToFreeImageFormat(),
                    GetSaveFlags()
                );
            }
            catch (Exception e)
            {
                Logger.Fatal("ExportViaFreeImage: FreeImageBitmap.Save() threw an exception!");
                Logger.Fatal(e);
                throw;
            }
            finally
            {
                Cancelling -= Exporter_Cancelling;
                PercentCompleted = 50;
            }
        }

        private void ConvertColor(FreeImageBitmap freeImageBitmap)
        {
            if (Preset.UseColorProfile)
            {
                Logger.Info("ConvertColor: Convert color using profile");
                ViewModels.ColorProfileViewModel targetProfile =
                    ViewModels.ColorProfileViewModel.CreateFromName(Preset.ColorProfile);
                targetProfile.TransformFromStandardProfile(freeImageBitmap);
                freeImageBitmap.ConvertColorDepth(Preset.ColorSpace.ToFreeImageColorDepth());
            }
            else
            {
                Logger.Info("ConvertColor: Convert color without profile");
                freeImageBitmap.ConvertColorDepth(Preset.ColorSpace.ToFreeImageColorDepth());
            }
            if (Preset.ColorSpace == ColorSpace.Monochrome)
            {
                SetMonochromePalette(freeImageBitmap);
            }
        }

        private void SetMonochromePalette(FreeImageBitmap freeImageBitmap)
        {
            Logger.Info("SetMonochromePalette: Convert to monochrome");
            freeImageBitmap.Palette.SetValue(new RGBQUAD(Color.Black), 0);
            freeImageBitmap.Palette.SetValue(new RGBQUAD(Color.White), 1);
        }

        private void ExportEmf(Metafile metafile)
        {
            Logger.Info("ExportEmf: exporting...");
            IntPtr handle = metafile.GetHenhmetafile();
            PercentCompleted = 50;
            Logger.Info("ExportEmf, handle: {0}", handle);
            IntPtr copy = Bovender.Unmanaged.Pinvoke.CopyEnhMetaFile(handle, FileName);
            Logger.Info("Deleting meta file handle: {0}", copy);
            Bovender.Unmanaged.Pinvoke.DeleteEnhMetaFile(copy);
            PercentCompleted = 100;
        }

        private FREE_IMAGE_SAVE_FLAGS GetSaveFlags()
        {
            Logger.Info("GetSaveFlags: {0}", Preset.FileType);
            switch (Preset.FileType)
            {
                case FileType.Png:
                    return FREE_IMAGE_SAVE_FLAGS.PNG_Z_BEST_COMPRESSION |
                           FREE_IMAGE_SAVE_FLAGS.PNG_INTERLACED;
                case FileType.Tiff:
                    switch (Preset.ColorSpace)
                    {
                        case ColorSpace.Monochrome:
                            return FREE_IMAGE_SAVE_FLAGS.TIFF_CCITTFAX4;
                        case ColorSpace.Cmyk:
                            return FREE_IMAGE_SAVE_FLAGS.TIFF_CMYK | FREE_IMAGE_SAVE_FLAGS.TIFF_LZW;
                        default:
                            return FREE_IMAGE_SAVE_FLAGS.TIFF_LZW;
                    }
                default:
                    return FREE_IMAGE_SAVE_FLAGS.DEFAULT;
            }
        }

        private void Exporter_Cancelling(object sender, Bovender.Mvvm.Models.ProcessModelEventArgs args)
        {
            if (_tiledBitmap != null)
            {
                _tiledBitmap.Cancel();
            }
        }

        /// <summary>
        /// Determines the effective transparency, which takes the file type into account.
        /// </summary>
        private Transparency EffectiveTransparency()
        {
            Transparency t;
            if (Preset.FileType.SupportsTransparency() && Preset.ColorSpace.SupportsTransparency())
            {
                t = Preset.Transparency;
                Logger.Info("SetTransparency: Transparency supported, passing through {0}", t);
            }
            else
            {
                t = Transparency.WhiteCanvas;
                Logger.Info("SetTransparency: Transparency not supported, using {0}", t);
            }
            return t;
        }

        #endregion

        #region Protected properties

        protected SelectionViewModel SelectionViewModel
        {
            get
            {
                if (_selectionViewModel == null)
                {
                    _selectionViewModel = new SelectionViewModel(Instance.Default.Application);
                }
                return _selectionViewModel;
            }
        }

        #endregion

        #region Private fields

        private DllManager _dllManager;
        private SingleExportSettings _settings;
        private Preset _preset;
        private bool _disposed;
        private Dictionary<FileType, FREE_IMAGE_FORMAT> _fileTypeToFreeImage;
        private TiledBitmap _tiledBitmap;
        private int _percentCompleted;
        private SelectionViewModel _selectionViewModel;

        #endregion

        #region Private constants
        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

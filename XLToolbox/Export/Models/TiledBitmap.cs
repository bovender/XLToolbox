using FreeImageAPI;
/* Tile.cs
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
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace XLToolbox.Export.Models
{
    /// <summary>
    /// A tile of a bitmap.
    /// </summary>
    class TiledBitmap : IDisposable
    {
        #region Properties

        /// <summary>
        /// Gets the width of the bitmap tile.
        /// </summary>
        public int Width { get; private set; }

        /// <summary>
        /// Gets the actual height of the bitmap tile.
        /// </summary>
        public int Height { get; private set; }

        /// <summary>
        /// Gets the height that was originally desired
        /// when the object was constructed.
        /// </summary>
        public int OriginalHeight { get; private set; }

        /// <summary>
        /// Is true if the tile fits the entire image.
        /// </summary>
        public bool IsFullSize
        {
            get
            {
                return Height == OriginalHeight;
            }
        }

        #endregion

        #region Public methods

        public FreeImageBitmap CreateFreeImageBitmap(Metafile metafile, Transparency transparency)
        {
            if (IsFullSize)
            {
                return DrawAtOnce(metafile, transparency);
            }
            else
            {
                return DrawTiles(metafile, transparency);
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructs an instance given a desired width and height.
        /// If there is not enough memory for the entire bitmap,
        /// the height will be reduced as much as necessary.
        /// </summary>
        /// <param name="width">Width in pixels.</param>
        /// <param name="height">Desired height in pixels; the actual
        /// height may be reduced if tiling is necessary.</param>
        public TiledBitmap(int width, int height)
        {
            if (height <= 0)
	        {
		        throw new ArgumentOutOfRangeException("height", "Bitmap height must be greater than 0");
	        }
            if (width <= 0)
	        {
		        throw new ArgumentOutOfRangeException("width", "Bitmap width must be greater than 0");
	        }

            if (!CreateBitmap(width, height))
            {
                throw new TileException(
                    String.Format(
                        "Unable to create bitmap (desired: {0}x{1}; tried: {2}x{3})",
                        width, height, Width, Height)
                    );
            }
        }

        #endregion

        #region Private methods

        private bool CreateBitmap(int width, int height)
        {
            bool success = false;
            Width = width;
            OriginalHeight = height;
            Logger.Info("Attempting to create bitmap with {0}x{1} pixels.", width, height);
            // height /= 4; // only for testing
            while (!success && height > 1)
            {
                try
                {
                    _bitmap = new Bitmap(width, height);
                    Logger.Info("Created bitmap with {0}x{1} pixels.", width, height);
                    success = true;
                }
                catch (Exception)
                {
                    Logger.Info("Could not create bitmap with {0}x{1} pixels.", width, height);
                    height /= 2;
                }
            }
            Height = height;
            return success;
        }

        private FreeImageBitmap DrawAtOnce(Metafile metafile, Transparency transparency)
        {
            Logger.Info("Drawing bitmap at once.");
            Graphics g = CreateGraphics(transparency);
            g.DrawImage(metafile, 0, 0, Width, Height);

            if (transparency == Transparency.TransparentWhite)
            {
                _bitmap.MakeTransparent(Color.White);
            }

            Logger.Info("Creating FreeImage bitmap");
            FreeImageBitmap f = new FreeImageBitmap(_bitmap);
            g.Dispose();
            return f;
        }

        private FreeImageBitmap DrawTiles(Metafile metafile, Transparency transparency)
        {
            // http://stackoverflow.com/a/503201/270712 by rjmunro
            int numTiles = (OriginalHeight - 1) / Height + 1;
            Logger.Info("Preparing to draw {0} tiles...", numTiles);
            FreeImageBitmap f = new FreeImageBitmap(Width, OriginalHeight, PixelFormat.Format32bppArgb);
            Graphics g = CreateGraphics(transparency);
            int scanLineSize = Width * 4; // must align with 32 bits; with RGBA this is always the case
            Logger.Info("Scan line size: {0:0.0} kB", scanLineSize / 1024);

            int currentLine = 0;
            int metafileHeight = metafile.Height;
            int metafileWidth = metafile.Width;
            Rectangle destinationRect = new Rectangle(
                0, 0, Width, Height);
            while (currentLine < OriginalHeight)
            {
                Logger.Info("Drawing tile starting at y={0}...", currentLine);
                int currentTileHeight = Math.Min(Height, OriginalHeight - currentLine);
                g.Clear(Color.Transparent);
                Rectangle sourceRect = new Rectangle(
                    0, metafileHeight * currentLine / OriginalHeight,
                    metafileWidth, metafileHeight * Height / OriginalHeight);
                Logger.Debug("Source rect: {0}; destination rect: {1}", sourceRect, destinationRect);
                g.DrawImage(metafile, destinationRect, sourceRect, GraphicsUnit.Pixel);
                IntPtr scanlinePointer = f.GetScanlinePointer(currentLine);
                BitmapData bitmapData = _bitmap.LockBits(
                    new Rectangle(0, 0, Width, Height), 
                    ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                CopyMemory(scanlinePointer, bitmapData.Scan0,
                    Convert.ToUInt32(bitmapData.Stride * currentTileHeight));
                _bitmap.UnlockBits(bitmapData);
                currentLine += currentTileHeight;
            }
            g.Dispose();
            Logger.Info("Flipping FreeImage vertically.");
            f.RotateFlip(RotateFlipType.RotateNoneFlipY);
            return f;
        }

        private Graphics CreateGraphics(Transparency transparency)
        {
            Logger.Info("Creating graphics object.");
            Graphics g = Graphics.FromImage(_bitmap);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            g.FillRectangle(CreateBrush(transparency), 0, 0, Width, Height);
            return g;
        }

        private Brush CreateBrush(Transparency transparency)
        {
            if (transparency == Transparency.TransparentCanvas)
            {
                return Brushes.Transparent;
            }
            else
            {
                return Brushes.White;
            }
        }

        #endregion

        #region Private fields

        private Bitmap _bitmap;
        private bool _disposed;

        #endregion

        #region P/Invoke

        [DllImport("kernel32.dll", EntryPoint = "CopyMemory", SetLastError = false)]
        private static extern void CopyMemory(IntPtr dest, IntPtr src, uint count);

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion

        #region Disposal

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool calledFromDispose)
        {
            if (!_disposed)
            {
                // Free unmanaged resources (if any)
                // ...
                if (calledFromDispose)
                {
                    // Free managed resources...
                    _bitmap.Dispose();
                }
                _disposed = true;
            }
        }

        #endregion
    }
}

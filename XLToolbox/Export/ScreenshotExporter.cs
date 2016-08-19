/* ScreenShotExporter.cs
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
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using FreeImageAPI;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Unmanaged;
using System.Drawing.Imaging;
using Bovender.Extensions;

namespace XLToolbox.Export
{
    /// <summary>
    /// Exports graphical data in screenshot quality to a file.
    /// </summary>
    public class ScreenshotExporter
    {
        #region Public methods

        public void ExportSelection(string fileName)
        {
            FreeImageBitmap fi;
            if (Instance.Default.Application.Selection is Microsoft.Office.Interop.Excel.Range)
            {
                Logger.Info("ExportSelection: Exporting via DIB");
                fi = CreateFreeImageViaDIB();
            }
            else
            {
                Logger.Info("ExportSelection: Exporting directly");
                fi = CreateFreeImageDirectly();
            }
            
            Logger.Info("ExportSelection: Save to file");
            fi.SetResolution(102.42f, 102.42f);
            fi.Save(fileName,
                FREE_IMAGE_FORMAT.FIF_PNG,
                FREE_IMAGE_SAVE_FLAGS.PNG_Z_BEST_COMPRESSION | FREE_IMAGE_SAVE_FLAGS.PNG_INTERLACED);
            fi.ReleaseComObject();
        }

        #endregion

        #region Private methods

        private FreeImageBitmap CreateFreeImageDirectly()
        {
            FreeImageBitmap fi;
            using (DllManager dllManager = new DllManager())
            {
                dllManager.LoadDll("FreeImage.DLL");
                Instance.Default.Application.Selection.Copy();
                MemoryStream data = Clipboard.GetData("PNG") as MemoryStream;
                Logger.Info("CreateFreeImageDirectly: Create FreeImage bitmap");
                fi = FreeImageBitmap.FromStream(data);
            }
            return fi;
        }

        private FreeImageBitmap CreateFreeImageViaDIB()
        {
            FreeImageBitmap fi;
            using (DllManager dllManager = new DllManager())
            {
                dllManager.LoadDll("FreeImage.DLL");
                Logger.Info("CreateFreeImageViaDIB: Copy to clipboard and get data from it");
                Instance.Default.Application.Selection.Copy();
                MemoryStream stream = Clipboard.GetData(System.Windows.DataFormats.Dib) as MemoryStream;
                using (DibBitmap dibBitmap = new DibBitmap(stream))
                {
                    Logger.Info("CreateFreeImageViaDIB: Create FreeImage bitmap");
                    fi = new FreeImageBitmap(dibBitmap.Bitmap);
                    bool convertType = fi.ConvertType(FREE_IMAGE_TYPE.FIT_BITMAP, true);
                    Logger.Debug("CreateFreeImageViaDIB: FreeImageBitmap.ConvertType returned {0}", convertType);
                    bool convertColorDepth = fi.ConvertColorDepth(FREE_IMAGE_COLOR_DEPTH.FICD_24_BPP); // won't work with 32 bpp!
                    Logger.Debug("CreateFreeImageViaDIB: FreeImageBitmap.ConvertColorDepth returned {0}", convertColorDepth);
                    fi.RotateFlip(System.Drawing.RotateFlipType.RotateNoneFlipY);
                }
            }
            return fi;
        }

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

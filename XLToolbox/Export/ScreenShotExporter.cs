/* ScreenShotExporter.cs
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
using System.Windows;
using XLToolbox.Excel.Instance;
using FreeImageAPI;
using System.IO;
using Bovender.Unmanaged;

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
            using (DllManager dllManager = new DllManager())
            {
                dllManager.LoadDll("FreeImage.DLL");
                ExcelInstance.Application.Selection.Copy();
                MemoryStream data = Clipboard.GetData("PNG") as MemoryStream;
                FreeImageBitmap fi = FreeImageBitmap.FromStream(data);
                fi.SetResolution(102.42f, 102.42f);
                fi.Save(fileName,
                    FREE_IMAGE_FORMAT.FIF_PNG,
                    FREE_IMAGE_SAVE_FLAGS.PNG_Z_BEST_COMPRESSION);
            }
        }

        #endregion
    }
}

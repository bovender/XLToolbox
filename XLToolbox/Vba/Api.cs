/* Api.cs
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
using System.Runtime.InteropServices;
using XLToolbox.Export.Models;
using XLToolbox.Export.ViewModels;

namespace XLToolbox.Vba
{
    /// <summary>
    /// Exposes some of the XL Toolbox features to external programs, such
    /// as VBA code.
    /// </summary>
    /// <example>
    /// <code>
    /// Private Sub ExportSelectionUsingXLToolboxNG()
    ///     Dim addin As Office.COMAddIn
    ///     Dim apiObject As Object
    ///     Set addin = Application.COMAddIns("XLToolboxForExcel")
    ///     Set apiObject = addin.Object
    ///     Debug.Print apiObject.ExportSelection("test.png", 300, "gray", "white")
    /// End Sub
    /// </code>
    /// </example>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("8DDC0086-3BAB-4D31-B5FD-6DEE3A1C78C9")]
    public class Api
    {
        #region Singleton factory

        static Api _api;

        public static Api Default
        {
            get
            {
                if (_api == null)
                {
                    _api = new Api();
                }
                return _api;
            }
        }

        #endregion

        #region API methods

        /// <summary>
        /// Exports the current selection to a graphic file.
        /// </summary>
        /// <param name="fileName">File name to export to. The file extension
        /// must be one of '.tif', '.tiff', '.png' or '.emf'.</param>
        /// <param name="dpi">Resolution in dots per inch (DPI).</param>
        /// <param name="colorSpace">Target color space (one of 'RGB', 'GRAY',
        /// 'MONO'; if this string is empty, 'RGB' will be assumed).
        /// Note: CMYK is currently not supported by this API function.
        /// </param>
        /// <param name="transparency">Transparency (one of 'NONE', 'CANVAS', 'WHITE'; 
        /// if this string is empty, 'NONE' will be assumed).
        /// </param>
        /// <returns>
        /// 0 if successful;
        /// 1 if unknown file type;
        /// 2 if unknown color space;
        /// 3 if unknown transparency type.
        /// </returns>
        public int ExportSelection(
            string fileName,
            int dpi,
            string colorSpace,
            string transparency)
        {
            Logger.Info("Export selection");
            string ext = System.IO.Path.GetExtension(fileName).ToUpper();

            FileType ft;
            switch (ext)
            {
                case ".TIF": ft = FileType.Tiff; break;
                case ".TIFF": ft = FileType.Tiff; break;
                case ".PNG": ft = FileType.Png; break;
                case ".EMF": ft = FileType.Emf; break;
                default: return 1;
            }

            ColorSpace cs;
            switch (colorSpace.ToUpper())
            {
                case "": cs = ColorSpace.Rgb; break;
                case "RGB": cs = ColorSpace.Rgb; break;
                case "GRAY": cs = ColorSpace.GrayScale; break;
                case "MONO": cs = ColorSpace.Monochrome; break;
                default: return 2;
            }

            Transparency t;
            switch (transparency.ToUpper())
            {
                case "": t = Transparency.TransparentCanvas; break;
                case "NONE": t = Transparency.WhiteCanvas; break;
                case "CANVAS": t = Transparency.TransparentCanvas; break;
                case "WHITE": t = Transparency.TransparentWhite; break;
                default: return 3;
            }

            Preset preset = new Preset(ft, dpi, cs);
            preset.Transparency = t;
            Logger.Info("Preset: {0}", preset);
            SingleExportSettings settings = SingleExportSettings.CreateForSelection(preset);
            SingleExportSettingsViewModel vm = new SingleExportSettingsViewModel(settings);

            vm.FileName = fileName;
            vm.ExportCommand.Execute(null);
            return 0; // success
        }

        /// <summary>
        /// Executes an XL Toolbox command. This method exists to facilitate
        /// using Application.OnKey which expects an Excel macro of VBA
        /// method as parameter, but does not work with .NET code. Of course,
        /// it can also be used to trigger XL Toolbox commands from independent
        /// VBA code.
        /// </summary>
        /// <param name="command">XL Toolbox command to execute</param>
        public void Execute(string command)
        {
            Logger.Info("Executing '{0}'", command);
            Command c;
            if (Enum.TryParse<Command>(command, out c))
            {
                Dispatcher.Execute(c);
            }
            else
            {
                Logger.Fatal("Parse failure: unknown command");
                throw new ArgumentException("Unknown command");
            }
        }

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

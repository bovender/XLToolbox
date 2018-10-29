/* Api.cs
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
using System.Runtime.InteropServices;
using Bovender.Extensions;
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
    [ClassInterface(ClassInterfaceType.None)]
    public class Api : IApi
    {
        #region Singleton factory

        public static Api Default
        {
            get
            {
                return _lazy.Value;
            }
        }

        private static readonly Lazy<Api> _lazy = new Lazy<Api>(() =>
        {
            Logger.Info("Api: Creating instance");
            return new Api();
        });

        #endregion

        #region Public properties

        public string LastLog { get; protected set; }

        #endregion

        #region API methods

        /// <summary>
        /// Returns the current XL Toolbox version as a string.
        /// </summary>
        /// <remarks>
        /// The XL Toolbox follows the semantic versioning scheme,
        /// see http://semver.org
        /// </remarks>
        /// <returns>Current XL Toolbox version</returns>
        public string Version()
        {
            return Versioning.SemanticVersion.Current.ToString();
        }

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

        /// <summary>
        /// Shows an exception.
        /// </summary>
        /// <param name="message">Exception message.</param>
        public void ShowException(string message)
        {
            VbaException e = new VbaException(message);
            Logger.Warn("VBA code called the XLToolbox.Vba.Api.Throw method", e);
            VbaExceptionViewModel vm = new VbaExceptionViewModel(e);
            vm.InjectInto<VbaExceptionView>().ShowDialogInForm();
        }

        /// <summary>
        /// Write a message to the log.
        /// </summary>
        public void Log(string message)
        {
            LastLog = message;
            Logger.Info(String.Format("Log(\"{0}\")", message));
        }

        /// <summary>
        /// Returns true if running in debug mode.
        /// </summary>
        public bool IsDebugMode()
        {
            #if DEBUG
                return true;
            #else
                return false;
            #endif
        }

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

/* LegacyToolbox.cs
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
using System.IO;
using System.Linq;
using System.Text;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Legacy
{
    class LegacyToolbox
    {
        private const string ADDIN_RESOURCE_NAME = "XLToolbox.Legacy.xltb-legacy-addin.xlam";

        #region Singleton

        public static LegacyToolbox Default
        {
            get { return _lazy.Value; }
        }

        #endregion

        #region Constructor

        private LegacyToolbox() { }

        #endregion

        #region Public command method

        public void RunCommand(Command command)
        {
            Microsoft.Office.Interop.Excel.Application app = Instance.Default.Application;
            switch (command)
            {
                case Command.OpenFromCell:
                    app.Run("RunOpenFromCell");
                    break;
                case Command.CopyPageSetup:
                    app.Run("RunCopyPageSetup");
                    break;
                case Command.SelectAllShapes:
                    app.Run("RunSelectAllShapes");
                    break;
                case Command.Anova1Way:
                    app.Run("RunANOVA");
                    break;
                case Command.Anova2Way:
                    app.Run("RunTwoWayANOVA");
                    break;
                case Command.FormulaBuilder:
                    app.Run("RunFormulaBuilder");
                    break;
                case Command.SelectionAssistant:
                    app.Run("RunSelectionAssistant");
                    break;
                case Command.LinearRegression:
                    app.Run("RunLinearRegression");
                    break;
                case Command.Correlation:
                    app.Run("RunCorrelation");
                    break;
                case Command.TransposeWizard:
                    app.Run("RunTransposeWizard");
                    break;
                case Command.MultiHisto:
                    app.Run("RunMultiHistogram");
                    break;
                case Command.Allocate:
                    app.Run("RunGroupAllocation");
                    break;
                case Command.AutomaticErrorBars:
                    app.Run("RunErrorBarsAuto");
                    break;
                case Command.InteractiveErrorBars:
                    app.Run("RunErrorBarsInteractive");
                    break;
                case Command.ChartDesign:
                    app.Run("RunChartDesign");
                    break;
                case Command.MoveDataSeriesLeft:
                    app.Run("RunMoveDataSeriesLeft");
                    break;
                case Command.MoveDataSeriesRight:
                    app.Run("RunMoveDataSeriesRight");
                    break;
                case Command.Annotate:
                    app.Run("RunChartAnnotation");
                    break;
                case Command.SpreadScatter:
                    app.Run("RunSpreadScatter");
                    break;
                case Command.SeriesToFront:
                    app.Run("RunSeriesToFront");
                    break;
                case Command.SeriesForward:
                    app.Run("RunSeriesForward");
                    break;
                case Command.SeriesBackward:
                    app.Run("RunSeriesBackward");
                    break;
                case Command.SeriesToBack:
                    app.Run("RunSeriesToBack");
                    break;
                case Command.AddSeries:
                    app.Run("RunAddSeries");
                    break;
                case Command.CopyChart:
                    app.Run("RunCopyChart");
                    break;
                case Command.PointChart:
                    app.Run("RunPointChart");
                    break;
                case Command.Watermark:
                    app.Run("RuneWatermark");
                    break;
                case Command.Prefs:
                    app.Run("RunPreferences");
                    break;
                default:
                    throw new InvalidOperationException("Unknown legacy command " + command.ToString());
            }
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Loads the legacy .xlam file if it has not been loaded yet.
        /// </summary>
        private static void LoadAddinIfNeeded()
        {
            Stream resourceStream = typeof(LegacyToolbox).Assembly
                .GetManifestResourceStream(ADDIN_RESOURCE_NAME);
            if (resourceStream == null)
            {
                throw new IOException("Unable to open resource stream " + ADDIN_RESOURCE_NAME);
            }
            string tempDir = Path.GetTempPath();
            string addinFile = Path.Combine(tempDir, ADDIN_RESOURCE_NAME);
            Stream tempStream = File.Create(addinFile);
            resourceStream.CopyTo(tempStream);
            tempStream.Close();
            resourceStream.Close();
            Instance.Default.Application.Workbooks.Open(addinFile);
        }

        #endregion

        #region Private static fields

        private static Lazy<LegacyToolbox> _lazy = new Lazy<LegacyToolbox>(
            () =>
            {
                LoadAddinIfNeeded();
                return new LegacyToolbox();
            }
        );

        #endregion
    }
}

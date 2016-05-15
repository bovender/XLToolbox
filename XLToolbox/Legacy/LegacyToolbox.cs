using Bovender.Mvvm.Actions;
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
using System.IO;
using System.Threading.Tasks;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Legacy
{
    public class LegacyToolbox
    {
        private const string ADDIN_RESOURCE_NAME = "XLToolbox.Legacy.xltb-legacy-addin.xlam";

        #region Static methods

        /// <summary>
        /// Detects if a .XLAM or .XLA version of the XL Toolbox is installed,
        /// and if so, issues a notice to the user and deactivates the VBA add-in.
        /// </summary>
        public static void DeactivateObsoleteVbaAddin()
        {
            foreach (Microsoft.Office.Interop.Excel.AddIn addin in Instance.Default.Application.AddIns)
            {
                if (addin.Name.StartsWith("Daniels XL Toolbox.xla") && addin.Installed)
                {
                    NotificationAction a = new NotificationAction();
                    a.Message = Strings.LegacyToolboxDetectedMessage;
                    a.Caption = Strings.LegacyToolboxDetectedCaption;
                    a.OkButtonLabel = Strings.OK;
                    a.Invoke();
                    addin.Installed = false;
                    break;
                }
            }
        }

        #endregion

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
                    app.Run("RunANOVA2Way");
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

        #region Private static fields and properties

        private static Lazy<LegacyToolbox> _lazy = new Lazy<LegacyToolbox>(
            () =>
            {
                Logger.Info("Initializing LegacyToolbox singleton");
                Instance.Default.LoadAddinFromEmbeddedResource(ADDIN_RESOURCE_NAME);
                return new LegacyToolbox();
            }
        );

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

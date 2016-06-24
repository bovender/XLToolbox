/* Dispatcher.cs
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
using System.Windows;
using Bovender.Extensions;
using Bovender.Mvvm.Actions;
using Bovender.Mvvm.Messaging;
using XLToolbox.ExceptionHandler;
using XLToolbox.Excel.ViewModels;
using XLToolbox.About;
using XLToolbox.Versioning;
using XLToolbox.SheetManager;
using XLToolbox.Export.ViewModels;
using Xl = Microsoft.Office.Interop.Excel;
using XLToolbox.Export.Models;
using Bovender.Mvvm.Models;

namespace XLToolbox
{
    /// <summary>
    /// Central dispatcher for all UI-initiated XL Toolbox commands.
    /// </summary>
    public static class Dispatcher
    {
        #region Public method

        /// <summary>
        /// Central command dispatcher. This public method also contains
        /// the central error handler for user-friendly error messages.
        /// </summary>
        /// <remarks>
        /// An enum-based approach was choosen in favor of publicly
        /// accessible static methods to enable listing of commands,
        /// e.g. for key bindings.
        /// </remarks>
        /// <param name="cmd">XL Toolbox command to execute.</param>
        public static void Execute(Command cmd)
        {
            Logger.Info("*** Execute {0} ***", cmd);
            try
            {
                switch (cmd)
                {
                    case Command.About: About(); break;
                    case Command.CheckForUpdates: CheckForUpdates(); break;
                    case Command.SheetManager: SheetManager(); break;
                    case Command.ExportSelection: ExportSelection(); break;
                    case Command.ExportSelectionLast: ExportSelectionLast(); break;
                    case Command.BatchExport: BatchExport(); break;
                    case Command.BatchExportLast: BatchExportLast(); break;
                    case Command.ExportScreenshot: ExportScreenshot(); break;
                    case Command.Donate: OpenDonatePage(); break;
                    case Command.ThrowError: throw new ExceptionHandler.TestException("This exception was thrown for testing purposes");
                    case Command.QuitExcel: QuitExcel(); break;
                    case Command.OpenCsv: OpenCsv(); break;
                    case Command.OpenCsvWithParams: OpenCsvWithSettings(); break;
                    case Command.SaveCsv: SaveCsv(); break;
                    case Command.SaveCsvWithParams: SaveCsvWithSettings(); break;
                    case Command.SaveCsvRange: SaveCsvRange(); break;
                    case Command.SaveCsvRangeWithParams: SaveCsvRangeWithSettings(); break;
                    case Command.Anova1Way: Anova1Way(); break;
                    case Command.Anova2Way: Anova2Way(); break;
                    case Command.AnovaRepeat: LastAnova(); break;
                    case Command.AutomaticErrorBars: ErrorBarsAutomatic(); break;
                    case Command.InteractiveErrorBars: ErrorBarsInteractive(); break;
                    case Command.LastErrorBars: LastErrorBars(); break;
                    case Command.UserSettings: EditUserSettings(); break;
                    case Command.OpenFromCell:
                    case Command.CopyPageSetup:
                    case Command.SelectAllShapes:
                    case Command.FormulaBuilder:
                    case Command.SelectionAssistant:
                    case Command.LinearRegression:
                    case Command.Correlation:
                    case Command.TransposeWizard:
                    case Command.MultiHisto:
                    case Command.Allocate:
                    case Command.ChartDesign:
                    case Command.MoveDataSeriesLeft:
                    case Command.MoveDataSeriesRight:
                    case Command.Annotate:
                    case Command.SpreadScatter:
                    case Command.SeriesToFront:
                    case Command.SeriesForward:
                    case Command.SeriesBackward:
                    case Command.SeriesToBack:
                    case Command.AddSeries:
                    case Command.CopyChart:
                    case Command.PointChart:
                    case Command.Watermark:
                    case Command.LegacyPrefs:
                        Legacy.LegacyToolbox.Default.RunCommand(cmd); break;
                    case Command.Shortcuts: EditShortcuts(); break;
                    case Command.SaveAs: SaveAs(); break;
                    default:
                        Logger.Fatal("No case has been implemented yet for this command");
                        throw new NotImplementedException("Don't know what to do with " + cmd.ToString());
                }
            }
            catch (Exception e)
            {
                Logger.Fatal(e, "Dispatcher exception");
                ExceptionViewModel vm = new ExceptionViewModel(e);
                vm.InjectInto<ExceptionView>().ShowDialogInForm();
            }
        }

        #endregion

        #region Private dispatching methods

        static void EditUserSettings()
        {
            UserSettings.UserSettingsViewModel vm = new UserSettings.UserSettingsViewModel();
            vm.InjectInto<UserSettings.UserSettingsView>().ShowDialogInForm();
        }

        static void About()
        {
            AboutViewModel avm = new AboutViewModel();
            Window w = avm.InjectInto<AboutView>();
            w.ShowDialogInForm();
        }

        static void CheckForUpdates()
        {
            EventHandler<MessageArgs<ProcessMessageContent>> h =
                (object sender, MessageArgs<ProcessMessageContent> args) =>
                {
                    args.Content.Caption = Strings.CheckingForUpdates;
                    Window view = args.Content.InjectInto<UpdaterProcessView>();
                    args.Content.ViewModel.ViewDispatcher = view.Dispatcher;
                    view.Show();
                };
            Versioning.UpdaterViewModel.Instance.CheckForUpdateMessage.Sent += h;
            Versioning.UpdaterViewModel.Instance.CheckForUpdateCommand.Execute(null);
            Versioning.UpdaterViewModel.Instance.CheckForUpdateMessage.Sent -= h;
        }

        static void SheetManager()
        {
            SheetManagerPane.Default.Visible = true;
            // wvm.InjectAndShowInThread<WorkbookView>();
        }

        static void ExportSelection()
        {
            Preset preset = UserSettings.UserSettings.Default.ExportPreset;
            if (preset == null)
            {
                preset = PresetsRepository.Default.First;
            }
            SingleExportSettings settings = SingleExportSettings.CreateForSelection(preset);
            SingleExportSettingsViewModel vm = new SingleExportSettingsViewModel(settings);
            vm.ShowProgressMessage.Sent += ShowProgressMessage_Sent;
            vm.ProcessFailedMessage.Sent += ProcessFailedMessage_Sent;
            vm.InjectInto<Export.Views.SingleExportSettingsView>().ShowDialogInForm();
        }

        static void ExportSelectionLast()
        {
            Export.QuickExporter quickExporter = new Export.QuickExporter();
            quickExporter.ExportSelection();
        }

        static void BatchExport()
        {
            BatchExportSettingsViewModel vm = BatchExportSettingsViewModel.FromLastUsed(
                Instance.Default.ActiveWorkbook);
            if (vm == null)
            {
                vm = new BatchExportSettingsViewModel();
            }
            vm.SanitizeOptions();
            vm.ShowProgressMessage.Sent += ShowProgressMessage_Sent;
            vm.ProcessFailedMessage.Sent += ProcessFailedMessage_Sent;
            vm.InjectInto<Export.Views.BatchExportSettingsView>().ShowDialogInForm();
        }

        static void BatchExportLast()
        {
            Export.QuickExporter quickExporter = new Export.QuickExporter();
            quickExporter.ExportBatch();
        }

        static void ExportScreenshot()
        {
            ScreenshotExporterViewModel vm = new ScreenshotExporterViewModel();
            if (vm.ExportSelectionCommand.CanExecute(null))
            {
                vm.ChooseFileNameMessage.Sent += (sender, args) =>
                    {
                        ChooseFileSaveAction a = new ChooseFileSaveAction();
                        a.Invoke(args);
                    };
                vm.ExportSelectionCommand.Execute(null);
            }
            else
            {
                NotificationAction a = new NotificationAction();
                a.Caption = Strings.ScreenshotExport;
                a.Message = Strings.ScreenshotExportRequiresGraphic;
                a.Invoke();
            }
        }

        static void OpenDonatePage()
        {
            System.Diagnostics.Process.Start(Properties.Settings.Default.DonateUrl);
        }

        static void QuitExcel()
        {
            if (Instance.Default.CountOpenWorkbooks > 1 || Instance.Default.CountUnsavedWorkbooks > 0)
            {
                Instance.Default.InjectInto<Excel.Views.QuitView>().ShowDialogInForm();
            }
            else
            {
                Instance.Default.Dispose();
            }
        }

        static void OpenCsv()
        {
            Csv.CsvImportViewModel vm = Csv.CsvImportViewModel.FromLastUsed();
            vm.ChooseImportFileNameMessage.Sent += (sender, args) =>
            {
                ChooseFileOpenAction a = new ChooseFileOpenAction();
                a.Invoke(args);
            };
            vm.ChooseFileNameCommand.Execute(null);
        }

        static void OpenCsvWithSettings()
        {
            Csv.CsvImportViewModel.FromLastUsed().InjectInto<Csv.CsvFileView>().ShowDialogInForm();
        }

        static void SaveCsv()
        {
            SaveCsv(null);
        }

        static void SaveCsv(Xl.Range range)
        {
            Csv.CsvExportViewModel vm = CreateCsvExportViewModel(range);
            vm.ChooseExportFileNameMessage.Sent += (sender, args) =>
            {
                ChooseFileSaveAction a = new ChooseFileSaveAction();
                a.Invoke(args);
            };
            vm.ChooseFileNameCommand.Execute(null);
        }

        static void SaveCsvWithSettings()
        {
            SaveCsvWithSettings(null);
        }

        static void SaveCsvWithSettings(Xl.Range range)
        {
            CreateCsvExportViewModel(range).InjectInto<Csv.CsvFileView>().ShowDialogInForm();
        }

        static void SaveCsvRange()
        {
            if (CheckSelectionIsRange())
            {
                SaveCsv(Instance.Default.Application.Selection as Xl.Range);
            }
        }

        static void SaveCsvRangeWithSettings()
        {
            if (CheckSelectionIsRange())
            {
                SaveCsvWithSettings(Instance.Default.Application.Selection as Xl.Range);
            }
        }

        static void Anova1Way()
        {
            UserSettings.UserSettings.Default.LastAnova = 1;
            Legacy.LegacyToolbox.Default.RunCommand(Command.Anova1Way);
        }

        static void Anova2Way()
        {
            UserSettings.UserSettings.Default.LastAnova = 2;
            Legacy.LegacyToolbox.Default.RunCommand(Command.Anova2Way);
        }

        static void LastAnova()
        {
            Command c = UserSettings.UserSettings.Default.LastAnova == 2 ?
                Command.Anova2Way : Command.Anova1Way;
            Legacy.LegacyToolbox.Default.RunCommand(c);
        }

        static void ErrorBarsAutomatic()
        {
            UserSettings.UserSettings.Default.LastErrorBars = 1;
            Legacy.LegacyToolbox.Default.RunCommand(Command.AutomaticErrorBars);
        }

        static void ErrorBarsInteractive()
        {
            UserSettings.UserSettings.Default.LastErrorBars = 2;
            Legacy.LegacyToolbox.Default.RunCommand(Command.InteractiveErrorBars);
        }

        static void LastErrorBars()
        {
            Command c = UserSettings.UserSettings.Default.LastErrorBars == 2 ? 
                Command.InteractiveErrorBars : Command.AutomaticErrorBars;
            Legacy.LegacyToolbox.Default.RunCommand(c);
        }

        static void EditShortcuts()
        {
            Keyboard.ManagerViewModel vm = new Keyboard.ManagerViewModel();
            vm.InjectInto<Keyboard.ManagerView>().ShowDialogInForm();
        }

        static void SaveAs()
        {
            Xl.Workbook w = Instance.Default.ActiveWorkbook;
            if (w != null)
            {
                Instance.Default.Application.Dialogs[Xl.XlBuiltInDialog.xlDialogSaveAs].Show();
            }
        }

        #endregion

        #region Private helper methods

        static bool CheckSelectionIsRange()
        {
            Xl.Range range = Instance.Default.Application.Selection as Xl.Range;
            if (range == null)
            {
                NotificationAction a = new NotificationAction(
                    Strings.RangeSelectionRequired, Strings.ActionRequiresSelectionOfCells, Strings.OK);
                a.Invoke();
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Creates an instance of XLToolbox.Csv.CsvExportViewModel and wires up the
        /// message events to display a progress bar and error message as needed.
        /// </summary>
        static Csv.CsvExportViewModel CreateCsvExportViewModel(Xl.Range range)
        {
            Csv.CsvExportViewModel vm = Csv.CsvExportViewModel.FromLastUsed();
            vm.Range = range;
            vm.ShowProgressMessage.Sent += (sender, args) =>
            {
                args.Content.CancelButtonText = Strings.Cancel;
                args.Content.Caption = Strings.ExportCsvFile;
                args.Content.CompletedMessage.Sent += (sender2, args2) =>
                {
                    args.Content.CloseViewCommand.Execute(null);
                };
                args.Content.InjectInto<Bovender.Mvvm.Views.ProcessView>().Show();
            };
            vm.ProcessFailedMessage.Sent += (sender, args) =>
            {
                Logger.Info("Received ExportFailedMessage, informing user");
                Bovender.Mvvm.Actions.ProcessCompletedAction action = new ProcessCompletedAction(
                    args.Content, Strings.CsvExportFailed, Strings.CsvExportFailed, Strings.Close);
                action.Invoke(args);
            };
            return vm;
        }

        internal static void ShowProgressMessage_Sent(object sender, MessageArgs<ProcessMessageContent> e)
        {
            Logger.Info("Creating process view");
            e.Content.CancelButtonText = Strings.Cancel;
            e.Content.Caption = Strings.Export;
            e.Content.CompletedMessage.Sent += (sender2, args2) =>
            {
                e.Content.CloseViewCommand.Execute(null);
            };
            e.Content.InjectInto<Bovender.Mvvm.Views.ProcessView>().ShowDialogInForm();
        }

        internal static void ProcessFailedMessage_Sent(object sender, MessageArgs<ProcessMessageContent> e)
        {
            Logger.Info("Received ExportFailedMessage, informing user");
            Bovender.Mvvm.Actions.ProcessCompletedAction action = new ProcessCompletedAction(
                e.Content, Strings.ExportFailed, Strings.ExportFailedMessage, Strings.Close);
            action.Invoke(e);
        }

        #endregion

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

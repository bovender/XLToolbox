/* Dispatcher.cs
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
using System.Windows;
using Bovender.Mvvm;
using Bovender.Mvvm.Actions;
using Bovender.Mvvm.Messaging;
using XLToolbox.ExceptionHandler;
using XLToolbox.Excel.ViewModels;
using XLToolbox.About;
using XLToolbox.Versioning;
using XLToolbox.SheetManager;
using XLToolbox.Export.ViewModels;
using Xl = Microsoft.Office.Interop.Excel;

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
                    case Command.ThrowError: throw new InsufficientMemoryException();
                    case Command.QuitExcel: QuitExcel(); break;
                    case Command.OpenCsv: OpenCsv(); break;
                    case Command.OpenCsvWithParams: OpenCsvWithSettings(); break;
                    case Command.SaveCsv: SaveCsv(); break;
                    case Command.SaveCsvWithParams: SaveCsvWithSettings(); break;
                    case Command.SaveCsvRange: SaveCsvRange(); break;
                    case Command.SaveCsvRangeWithParams: SaveCsvRangeWithSettings(); break;
                    default:
                        throw new NotImplementedException("Don't know what to do with " + cmd.ToString());
                }
            }
            catch (Exception e)
            {
                ExceptionViewModel vm = new ExceptionViewModel(e);
                vm.InjectInto<ExceptionView>().ShowDialog();
            }
        }

        #endregion

        #region Private dispatching methods

        static void About()
        {
            AboutViewModel avm = new AboutViewModel();
            avm.InjectInto<AboutView>().ShowDialog();
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
            WorkbookViewModel wvm = new WorkbookViewModel(Instance.Default.ActiveWorkbook);
            Workarounds.ShowModelessInExcel<WorkbookView>(wvm);
        }

        static void ExportSelection()
        {
            SingleExportSettingsViewModel vm = new SingleExportSettingsViewModel();
            vm.InjectInto<Export.Views.SingleExportSettingsView>().ShowDialog();
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
            vm.InjectInto<Export.Views.BatchExportSettingsView>().ShowDialog();
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
            if (Instance.Default.CountOpenWorkbooks > 0)
            {
                Instance.Default.InjectInto<Excel.Views.QuitView>().ShowDialog();
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
            Csv.CsvImportViewModel.FromLastUsed().InjectInto<Csv.CsvFileView>().ShowDialog();
        }

        static void SaveCsv()
        {
            SaveCsv(null);
        }

        static void SaveCsv(Xl.Range range)
        {
            Csv.CsvExportViewModel vm = Csv.CsvExportViewModel.FromLastUsed();
            vm.Range = range;
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
            Csv.CsvExportViewModel vm = Csv.CsvExportViewModel.FromLastUsed();
            vm.Range = range;
            vm.InjectInto<Csv.CsvFileView>().ShowDialog();
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

        #endregion

        #region Private helper methods

        static bool CheckSelectionIsRange()
        {
            Xl.Range range = Instance.Default.Application.Selection as Xl.Range;
            if (range == null)
            {
                NotificationAction a = new NotificationAction();
                a.Caption = Strings.RangeSelectionRequired;
                a.Message = Strings.ActionRequiresSelectionOfCells;
                a.OkButtonLabel = Strings.OK;
                a.Invoke();
                return false;
            }
            else
            {
                return true;
            }
        }

        #endregion
    }
}

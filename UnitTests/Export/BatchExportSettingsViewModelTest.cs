using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using XLToolbox.Test;
using XLToolbox.Export.Models;
using XLToolbox.Export.ViewModels;
using XLToolbox.Excel.Instance;

namespace XLToolbox.Test.Export
{
    [TestFixture]
    class BatchExportSettingsViewModelTest
    {
        #region Command state tests
        [Test]
        public void ExecuteWithoutExcel()
        {
            BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();
            Assert.IsFalse(viewModel.ExportCommand.CanExecute(null));
        }

        [Test]
        public void ExecuteWithoutChart()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null));
            }
        }

        [Test]
        public void ExecuteWithOneChart()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithOneChartAndOneShape()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithTwoChartsAndOneShape()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 2);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithOneChartAndTwoShapes()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 2);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithOneShapeAndNoChart()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                // Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithOneChartOtherSheet()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                // Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                excel.App.Worksheets.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithTwoChartsOtherSheet()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 2);
                // Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                excel.App.Worksheets.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithTwoChartsOneShapeOtherSheet()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 2);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                excel.App.Worksheets.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithOneChartTwoShapesOtherSheet()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 2);
                excel.App.Worksheets.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithOneChartOtherWorkbook()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                // Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                excel.App.Workbooks.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Open workbook/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Open workbook/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Open workbook/charts and shapes/layout");
            }
        }

        [Test]
        public void ExecuteWithOneChartTwoShapesOtherWorkbook()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 2);
                excel.App.Workbooks.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active sheet/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Active workbook/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "All workbooks/charts/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Open workbook/charts and shapes/single items");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                    "Open workbook/charts/layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
                Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                    "Open workbook/charts and shapes/layout");
            }
        }

        #endregion

        #region Enum state tests

        [Test]
        public void StatesWithoutExcel()
        {
            BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();
            Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
            Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
            Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");
            Assert.IsFalse(viewModel.IsChartsEnabled, "Charts");
            Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Charts And Shapes");
            Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Preserve Layout");
        }

        [Test]
        public void StatesWithoutChart()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();
                Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");
                Assert.IsFalse(viewModel.IsChartsEnabled, "Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Charts And Shapes");
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Preserve Layout");
            }
        }

        [Test]
        public void StatesWithOneChart()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsTrue(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsTrue(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        [Test]
        public void StatesWithOneChartAndOneShape()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsTrue(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsTrue(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        [Test]
        public void StatesWithTwoChartsAndOneShape()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 2);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsTrue(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsTrue(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        [Test]
        public void StatesWithTwoShapesAndNoChart()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 2);
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsTrue(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsFalse(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }
    
        [Test]
        public void StatesWithChartsOnOtherSheet()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 2);
                excel.App.Worksheets.Add().Activate();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsTrue(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsFalse(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsTrue(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        [Test]
        public void StatesWithOneShapeOnOtherSheet()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                excel.App.Worksheets.Add().Activate();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsTrue(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsFalse(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        [Test]
        public void StatesWithOneChartAndOneShapeOnOtherSheet()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                excel.App.Worksheets.Add().Activate();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsTrue(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsFalse(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsTrue(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        [Test]
        public void StatesWithOneChartInOtherWorkbook()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                excel.App.Workbooks.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsTrue(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsFalse(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsTrue(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        [Test]
        public void StatesWithTwoChartsInOtherWorkbook()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                excel.App.Worksheets.Add().Activate();
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                excel.App.Workbooks.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsTrue(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsFalse(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsTrue(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        [Test]
        public void StatesWithChartsAndShapesInOtherWorkbook()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 1);
                excel.App.Worksheets.Add().Activate();
                Helpers.CreateSomeCharts(excel.App.ActiveSheet, 3);
                Helpers.CreateSomeShapes(excel.App.ActiveSheet, 1);
                excel.App.Workbooks.Add();
                BatchExportSettingsViewModel viewModel = new BatchExportSettingsViewModel();

                // Assert scope states
                Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
                Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
                Assert.IsTrue(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

                // Assert object states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                Assert.IsFalse(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                Assert.IsFalse(viewModel.IsChartsEnabled, "Active Workbook/Charts");
                Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                Assert.IsTrue(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
                Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

                // Assert layout states
                viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
                Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

                viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
                viewModel.Objects.AsEnum = BatchExportObjects.Charts;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
                Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
                viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
                Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
                Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
            }
        }

        #endregion
    }
}

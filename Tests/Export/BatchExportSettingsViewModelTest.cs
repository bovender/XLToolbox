/* BatchExportSettingsViewModelTest.cs
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
using System.Linq;
using System.Text;
using System.Threading;
using NUnit.Framework;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Test;
using XLToolbox.Export.Models;
using XLToolbox.Export.ViewModels;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Test.Export
{
    [TestFixture]
    class BatchExportSettingsViewModelTest
    {
        BatchExportSettings _settings;

        #region Setup

        [TestFixtureSetUp]
        public void TestFixtureSetup()
        {
            SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
        }

        [SetUp] 
        public void Setup()
        {
            Instance.Default.Reset();
            Instance.Default.CreateWorkbook();
            Preset preset = UserSettings.UserSettings.Default.ExportPresets.FirstOrDefault();
            if (preset == null)
            {
                preset = PresetsRepository.Default.First;
            }
            _settings = new BatchExportSettings(preset);
        }

        [TearDown]
        public void TearDown()
        {
        }

        #endregion

        #region Command state tests

        [Test]
        public void ExecuteWithoutExcel()
        {
            BatchExportSettingsViewModel viewModel = CreateVM();
            Assert.IsFalse(viewModel.ExportCommand.CanExecute(null));
        }

        [Test]
        public void ExecuteWithoutChart()
        {
            BatchExportSettingsViewModel viewModel = CreateVM();
            Assert.IsFalse(viewModel.ExportCommand.CanExecute(null));
        }

        [Test]
        public void ExecuteWithOneChart()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            BatchExportSettingsViewModel viewModel = CreateVM();

            viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                "Active sheet/charts/single items");

            viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                "Active workbook/charts/single items");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
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

        [Test]
        public void ExecuteWithOneChartAndOneShape()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            BatchExportSettingsViewModel viewModel = CreateVM();

            viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                "Active sheet/charts/single items");

            viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                "Active workbook/charts/single items");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
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

        [Test]
        public void ExecuteWithTwoChartsAndOneShape()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 2);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            BatchExportSettingsViewModel viewModel = CreateVM();

            viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                "Active sheet/charts/single items");

            viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                "Active workbook/charts/single items");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
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

        [Test]
        public void ExecuteWithOneChartAndTwoShapes()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 2);
            BatchExportSettingsViewModel viewModel = CreateVM();

            viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                "Active sheet/charts/single items");

            viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
                "Active workbook/charts/single items");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SingleItems;
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
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

        [Test]
        public void ExecuteWithOneShapeAndNoChart()
        {
            // Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            BatchExportSettingsViewModel viewModel = CreateVM();

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

        [Test]
        public void ExecuteWithOneChartOtherSheet()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            // Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Worksheets.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
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
            Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                "Active workbook/charts and shapes/layout");
        }

        [Test]
        public void ExecuteWithTwoChartsOtherSheet()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 2);
            // Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Worksheets.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
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

        [Test]
        public void ExecuteWithTwoChartsOneShapeOtherSheet()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 2);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Worksheets.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
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

        [Test]
        public void ExecuteWithOneChartTwoShapesOtherSheet()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 2);
            Instance.Default.Application.Worksheets.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.ExportCommand.CanExecute(null),
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

        [Test]
        public void ExecuteWithOneChartOtherWorkbook()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            // Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Workbooks.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
                "Open workbooks/charts and shapes/single items");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
            Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                "Open workbooks/charts/layout");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            viewModel.Layout.AsEnum = BatchExportLayout.SheetLayout;
            Assert.IsFalse(viewModel.ExportCommand.CanExecute(null),
                "Open workbooks/charts and shapes/layout");
        }

        [Test]
        public void ExecuteWithOneChartTwoShapesOtherWorkbook()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 2);
            Instance.Default.Application.Workbooks.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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

        #endregion

        #region Enum state tests

        [Test]
        public void StatesWithoutExcel()
        {
            BatchExportSettingsViewModel viewModel = CreateVM();
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
            BatchExportSettingsViewModel viewModel = CreateVM();
            Assert.IsFalse(viewModel.IsActiveSheetEnabled, "Active Sheet");
            Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
            Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");
            Assert.IsFalse(viewModel.IsChartsEnabled, "Charts");
            Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Charts And Shapes");
            Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Preserve Layout");
        }

        [Test]
        public void StatesWithOneChart()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            BatchExportSettingsViewModel viewModel = CreateVM();

            // Assert scope states
            Assert.IsTrue(viewModel.IsActiveSheetEnabled, "Active Sheet");
            Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
            Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

            // Assert object states
            viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
            Assert.IsTrue(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
            Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
            viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
            Assert.IsTrue(viewModel.IsChartsEnabled, "Active Workbook/Charts");
            Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            Assert.IsTrue(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
            Assert.IsFalse(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

            // Assert layout states
            viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "ActiveSheet/Charts and Shapes/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "ActiveSheet/Charts and Shapes/Preserve Layout");

            viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }

        [Test]
        public void StatesWithOneChartAndOneShape()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            BatchExportSettingsViewModel viewModel = CreateVM();

            // Assert scope states
            Assert.IsTrue(viewModel.IsActiveSheetEnabled, "Active Sheet");
            Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
            Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

            // Assert object states
            viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
            Assert.IsTrue(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
            Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
            viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
            Assert.IsTrue(viewModel.IsChartsEnabled, "Active Workbook/Charts");
            Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            Assert.IsTrue(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
            Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

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
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }

        [Test]
        public void StatesWithTwoChartsAndOneShape()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 2);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            BatchExportSettingsViewModel viewModel = CreateVM();

            // Assert scope states
            Assert.IsTrue(viewModel.IsActiveSheetEnabled, "Active Sheet");
            Assert.IsFalse(viewModel.IsActiveWorkbookEnabled, "All Sheets");
            Assert.IsFalse(viewModel.IsOpenWorkbooksEnabled, "All Workbooks");

            // Assert object states
            viewModel.Scope.AsEnum = BatchExportScope.ActiveSheet;
            Assert.IsTrue(viewModel.IsChartsEnabled, "ActiveSheet/Charts");
            Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "ActiveSheet/Charts And Shapes");
            viewModel.Scope.AsEnum = BatchExportScope.ActiveWorkbook;
            Assert.IsTrue(viewModel.IsChartsEnabled, "Active Workbook/Charts");
            Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            Assert.IsTrue(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
            Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

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
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }

        [Test]
        public void StatesWithTwoShapesAndNoChart()
        {
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 2);
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Active Workbook/Charts And Shapes");
            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            Assert.IsFalse(viewModel.IsChartsEnabled, "Open Workbooks/Charts");
            Assert.IsTrue(viewModel.IsChartsAndShapesEnabled, "Open Workbooks/Charts And Shapes");

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
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }
    
        [Test]
        public void StatesWithChartsOnOtherSheet()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 2);
            Instance.Default.Application.Worksheets.Add().Activate();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }

        [Test]
        public void StatesWithOneShapeOnOtherSheet()
        {
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Worksheets.Add().Activate();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsFalse(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }

        [Test]
        public void StatesWithOneChartAndOneShapeOnOtherSheet()
        {
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Worksheets.Add().Activate();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Active Workbook/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Active Workbook/Charts and Shapes/Preserve Layout");

            viewModel.Scope.AsEnum = BatchExportScope.OpenWorkbooks;
            viewModel.Objects.AsEnum = BatchExportObjects.Charts;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts/Preserve Layout");
            viewModel.Objects.AsEnum = BatchExportObjects.ChartsAndShapes;
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsTrue(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }

        [Test]
        public void StatesWithOneChartInOtherWorkbook()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Workbooks.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }

        [Test]
        public void StatesWithTwoChartsInOtherWorkbook()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Worksheets.Add().Activate();
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Workbooks.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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
            Assert.IsTrue(viewModel.IsSingleItemsEnabled, "Open Workbooks/Charts and Shapes/Single Items");
            Assert.IsFalse(viewModel.IsSheetLayoutEnabled, "Open Workbooks/Charts and Shapes/Preserve Layout");
        }

        [Test]
        public void StatesWithChartsAndShapesInOtherWorkbook()
        {
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Worksheets.Add().Activate();
            Helpers.CreateSomeCharts(Instance.Default.Application.ActiveSheet, 3);
            Helpers.CreateSomeShapes(Instance.Default.Application.ActiveSheet, 1);
            Instance.Default.Application.Workbooks.Add();
            BatchExportSettingsViewModel viewModel = CreateVM();

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

        private BatchExportSettingsViewModel CreateVM()
        {
            return new BatchExportSettingsViewModel(_settings);
        }

        #endregion
    }
}

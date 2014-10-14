using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Core;
using XLToolbox.Core.Excel;
using NUnit.Framework;

namespace XLToolbox.Test.Excel
{
    /// <summary>
    /// Unit tests for the XLToolbox.Core.Excel namespace.
    /// </summary>
    [TestFixture]
    class WorkbookViewModelTest
    {
        [SetUp]
        public static void SetUp()
        {
            ExcelInstance.Start();
        }

        [TearDown]
        public static void TearDown()
        {
            ExcelInstance.Shutdown();
        }

        [Test]
        public static void WorkbookViewModelProperties()
        {
            Workbook wb = ExcelInstance.Application.Workbooks.Add();
            WorkbookViewModel wvm = new WorkbookViewModel(wb);
            Assert.AreEqual(wvm.DisplayString, wb.Name,
                "WorkbookViewModel does not give workbook name as display string");

            // When accessing sheets in a collection, keep in mind that
            // the Sheets collection of a Workbook instance is 1-based.
            Assert.AreEqual(wvm.Sheets[0].DisplayString, wb.Sheets[1].Name,
                "SheetViewModel in WorkbookViewModel has incorrect sheet name");

            Assert.AreEqual(wvm.Sheets.Count, wb.Sheets.Count,
                "ViewModel and workbook report different sheet counts.");
        }

        [Test]
        public static void MoveSheetsUp()
        {
            Workbook wb = ExcelInstance.Application.Workbooks.Add();
            for (int i = 0; i < 2; i++)
            {
                wb.Sheets.Add();
            }

            WorkbookViewModel wvm = new WorkbookViewModel(wb);

            // Select the second sheet in the collection (index #1)
            SheetViewModel svm = wvm.Sheets[1];
            string sheetName = svm.DisplayString;

            // With no sheets selected, the move-up command should
            // be disabled.
            Assert.IsFalse(wvm.MoveSheetUp.CanExecute(null),
                "Move command is enabled, should be disabled with no sheets selected.");

            svm.IsSelected = true;
            Assert.IsTrue(wvm.MoveSheetUp.CanExecute(null),
                "Move command is disabled, should be enabled with one sheet selected.");
            wvm.MoveSheetUp.Execute(null);

            // The selected sheet should now be the first sheet,
            // which cannot be moved 'up' as it is at the 'top'
            // already; so the command should be disabled again.
            Assert.IsFalse(wvm.MoveSheetUp.CanExecute(null),
                "Move command is enabled, should be disabled if the very first sheet is selected.");

            // Check the the move was performed on the workbook too.
            Assert.AreEqual(sheetName, wb.Sheets[1].Name,
                "Moving the sheet was not performed on the actual workbook");
        }

        [Test]
        public static void MoveSheetsDown()
        {
            Workbook wb = ExcelInstance.Application.Workbooks.Add();
            for (int i = 0; i < 2; i++)
            {
                wb.Sheets.Add();
            }

            WorkbookViewModel wvm = new WorkbookViewModel(wb);

            // Select the second-to-last sheet in the collection
            SheetViewModel svm = wvm.Sheets[wvm.Sheets.Count - 2];
            string sheetName = svm.DisplayString;

            // With no sheets selected, the move-down command should
            // be disabled.
            Assert.IsFalse(wvm.MoveSheetDown.CanExecute(null),
                "Move-down command is enabled, should be disabled with no sheets selected.");

            svm.IsSelected = true;
            Assert.IsTrue(wvm.MoveSheetDown.CanExecute(null),
                "Move-down command is disabled, should be enabled with one sheet selected.");
            wvm.MoveSheetDown.Execute(null);

            // The selected sheet should now be the first sheet,
            // which cannot be moved 'up' as it is at the 'top'
            // already; so the command should be disabled again.
            Assert.IsFalse(wvm.MoveSheetDown.CanExecute(null),
                "Move-down command is enabled, should be disabled if the very last sheet is selected.");

            // Check the the move was performed on the workbook too.
            Assert.AreEqual(sheetName, wb.Sheets[wb.Sheets.Count].Name,
                "Moving the sheet down was not performed on the actual workbook");
        }

        [Test]
        public static void MoveSheetsToTop()
        {
            Workbook wb = ExcelInstance.Application.Workbooks.Add();
            for (int i = 0; i < 6; i++)
            {
                wb.Sheets.Add();
            }

            WorkbookViewModel wvm = new WorkbookViewModel(wb);
            
            // Without sheets selected, the Move-to-top command should be disabled
            Assert.IsFalse(wvm.MoveSheetsToTop.CanExecute(null),
                "The Move-to-top command should be disabled without selected sheets.");

            // Select the fourth and sixth sheets and remember their names
            SheetViewModel svm4 = wvm.Sheets[3];
            svm4.IsSelected = true;
            string sheetName4 = svm4.DisplayString;

            SheetViewModel svm6 = wvm.Sheets[5];
            svm6.IsSelected = true;
            string sheetName6 = svm6.DisplayString;

            // With sheets selected, the Move-to-top command should be disabled
            Assert.IsTrue(wvm.MoveSheetsToTop.CanExecute(null),
                "The Move-to-top command should be enabled with selected sheets.");

            wvm.MoveSheetsToTop.Execute(null);

            // Since a selected sheet was moved to the top, the command should
            // now be disabled again.
            Assert.IsFalse(wvm.MoveSheetsToTop.CanExecute(null),
                "The Move-to-top command should be disabled if the first sheet is selected.");

            // Verify that the display strings of the view models correspond to
            // the names of the worksheets in the workbook, to make sure that
            // the worksheets have indeed been rearranged as well.
            Assert.AreEqual(sheetName4, wb.Sheets[1].Name,
                "Moving the sheets to top was not performed on the actual workbook");
            Assert.AreEqual(sheetName6, wb.Sheets[2].Name,
                "Moving the sheets to top was not performed for all sheets on the actual workbook");
        }

        [Test]
        public static void MoveSheetsToBottom()
        {
            Workbook wb = ExcelInstance.Application.Workbooks.Add();
            for (int i = 0; i < 6; i++)
            {
                wb.Sheets.Add();
            }

            WorkbookViewModel wvm = new WorkbookViewModel(wb);

            // Without sheets selected, the Move-to-bottom command should be disabled
            Assert.IsFalse(wvm.MoveSheetsToBottom.CanExecute(null),
                "The Move-to-bottom command should be disabled without selected sheets.");

            // Select the fourth and sixth sheets and remember their names
            SheetViewModel svm2 = wvm.Sheets[1];
            svm2.IsSelected = true;
            string sheetName2 = svm2.DisplayString;

            SheetViewModel svm4 = wvm.Sheets[3];
            svm4.IsSelected = true;
            string sheetName4 = svm4.DisplayString;

            // With sheets selected, the Move-to-bottom command should be disabled
            Assert.IsTrue(wvm.MoveSheetsToBottom.CanExecute(null),
                "The Move-to-bottom command should be enabled with selected sheets.");

            wvm.MoveSheetsToBottom.Execute(null);

            // Since a selected sheet was moved to the bottom, the command should
            // now be disabled again.
            Assert.IsFalse(wvm.MoveSheetsToBottom.CanExecute(null),
                "The Move-to-Bottom command should be disabled if the last sheet is selected.");

            // Verify that the display strings of the view models correspond to
            // the names of the worksheets in the workbook, to make sure that
            // the worksheets have indeed been rearranged as well.
            Assert.AreEqual(sheetName2, wb.Sheets[wb.Sheets.Count-1].Name,
                "Moving the sheets to bottom was not performed on the actual workbook");
            Assert.AreEqual(sheetName4, wb.Sheets[wb.Sheets.Count].Name,
                "Moving the sheets to bottom was not performed for all sheets on the actual workbook");
        }

        [Test]
        public static void DeleteSheets()
        {
            Workbook wb = ExcelInstance.Application.Workbooks.Add();
            for (int i = 0; i < 6; i++)
            {
                wb.Sheets.Add();
            }

            int oldCount = wb.Sheets.Count;
            WorkbookViewModel wvm = new WorkbookViewModel(wb);
            Assert.IsFalse(wvm.DeleteSheets.CanExecute(null),
                "Delete sheets command should be disabled with no sheets selected.");
            wvm.Sheets[2].IsSelected = true;
            wvm.Sheets[4].IsSelected = true;
            int numSelected = wvm.NumSelectedSheets;
            Assert.IsTrue(wvm.DeleteSheets.CanExecute(null),
                "Delete sheets command should be enabled with some sheets selected.");
            wvm.DeleteSheets.Execute(null);

            wvm.DeleteSheets.Execute(null);
            Assert.AreEqual(oldCount - numSelected, wvm.Sheets.Count,
                "After deleting sheets, the workbook view model has unexpected number of sheet view models.");
            Assert.AreEqual(oldCount - numSelected, wb.Sheets.Count,
                "After deleting sheets, the workbook has unexpected number of sheets.");
        }
    }
}

/* WorkbookViewModelTest.cs
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
using Microsoft.Office.Interop.Excel;
using XLToolbox.Excel.ViewModels;
using NUnit.Framework;

namespace XLToolbox.Test.Excel
{
    /// <summary>
    /// Unit tests for the XLToolbox.Core.Excel namespace.
    /// </summary>
    [TestFixture]
    class WorkbookViewModelTest
    {
        Instance _excelInstance;

        [SetUp]
        public void SetUp()
        {
            _excelInstance = Instance.Default;
        }

        [TearDown]
        public void TearDown()
        {
            _excelInstance.Dispose();
        }

        [Test]
        public void WorkbookViewModelProperties()
        {
            Workbook wb = Instance.Default.CreateWorkbook();
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
        public void MoveSheetsUp()
        {
            Workbook wb = Instance.Default.CreateWorkbook(3);
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
        public void MoveSheetsDown()
        {
            Workbook wb = Instance.Default.CreateWorkbook(6);
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
        public void MoveSheetsToTop()
        {
            Workbook wb = Instance.Default.CreateWorkbook(8);
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
        public void MoveSheetsToBottom()
        {
            Workbook wb = Instance.Default.CreateWorkbook(8);
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
        public void DeleteSheets()
        {
            Workbook wb = Instance.Default.CreateWorkbook(8);
            int oldCount = wb.Sheets.Count;
            WorkbookViewModel wvm = new WorkbookViewModel(wb);
            Assert.IsFalse(wvm.DeleteSheets.CanExecute(null),
                "Delete sheets command should be disabled with no sheets selected.");
            wvm.Sheets[2].IsSelected = true;
            wvm.Sheets[4].IsSelected = true;
            string sheetName3 = wvm.Sheets[2].DisplayString;
            string sheetName5 = wvm.Sheets[4].DisplayString;
            int numSelected = wvm.NumSelectedSheets;
            Assert.IsTrue(wvm.DeleteSheets.CanExecute(null),
                "Delete sheets command should be enabled with some sheets selected.");

            bool messageSent = false;
            bool confirmDelete = false;
            wvm.ConfirmDeleteMessage.Sent += (sender, message) =>
            {
                messageSent = true;
                message.Content.Confirmed = confirmDelete;
                message.Respond();
            };

            wvm.DeleteSheets.Execute(null);
            Assert.IsTrue(messageSent, "No ViewModelMessage was sent.");
            Assert.AreEqual(oldCount, wvm.Sheets.Count,
                "Number of sheets was changed even though deletion was not confirmed.");

            confirmDelete = true;
            wvm.DeleteSheets.Execute(null);
            Assert.AreEqual(oldCount - numSelected, wvm.Sheets.Count,
                "After deleting sheets, the workbook view model has unexpected number of sheet view models.");
            Assert.AreEqual(oldCount - numSelected, wb.Sheets.Count,
                "After deleting sheets, the workbook has unexpected number of sheets.");
            object obj;
            Assert.Throws(typeof(System.Runtime.InteropServices.COMException), () =>
                {
                    obj = wb.Sheets[sheetName3];
                },
                String.Format("Sheet {0} (sheetName3) should have been deleted but is still there.", sheetName3)
            );
            Assert.Throws(typeof(System.Runtime.InteropServices.COMException), () =>
            {
                obj = wb.Sheets[sheetName5];
            },
                String.Format("Sheet {0} (sheetName5) should have been deleted but is still there.", sheetName5)
            );
        }

        [Test]
        public void SelectSheet()
        {
            Workbook wb = Instance.Default.CreateWorkbook(8);
            WorkbookViewModel wvm = new WorkbookViewModel(wb);
            wvm.Sheets[2].IsSelected = true;
            Assert.AreEqual(wvm.Sheets[2].DisplayString, wb.ActiveSheet.Name);
            wvm.Sheets[4].IsSelected = true;
            Assert.AreEqual(wvm.Sheets[4].DisplayString, wb.ActiveSheet.Name);
        }

        [Test]
        public void RenameSheet()
        {
            Workbook wb = Instance.Default.CreateWorkbook();
            WorkbookViewModel wvm = new WorkbookViewModel(wb);
            string oldName = wvm.Sheets[0].DisplayString;
            string newName = "valid new name";
            bool messageSent = false;
            Assert.False(wvm.RenameSheet.CanExecute(null),
                "Rename sheet command should be disabled with no sheet selected.");
            wvm.Sheets[0].IsSelected = true;
            Assert.True(wvm.RenameSheet.CanExecute(null),
                "Rename sheet command should be enabled with one sheet selected.");
            wvm.RenameSheetMessage.Sent += (sender, args) =>
                {
                    messageSent = true;
                    args.Content.Confirmed = true;
                    args.Content.Value = newName;
                    args.Respond();
                };
            wvm.RenameSheet.Execute(null);
            Assert.True(messageSent, "Rename sheet message was not sent.");
            Assert.AreEqual(newName, wvm.Sheets[0].DisplayString,
                String.Format(
                    "Worksheet name is '{0}', should have been renamed to '{1}'.",
                    wvm.Sheets[0].DisplayString, newName
                )
            );

            Assert.Throws<InvalidSheetNameException>(() =>
            {
                newName = "invalid\\sheet\\name";
                wvm.RenameSheet.Execute(null);
                Assert.Fail("Assigning an invalid name should cause an exception.");
            });
        }
    }
}

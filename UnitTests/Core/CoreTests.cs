using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Core;
using XLToolbox.Core.Excel;
using NUnit.Framework;

namespace XLToolbox.Test.Core
{
    [TestFixture]
    class CoreTests
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
        public static void TestSheetViewModel()
        {
            Worksheet ws = ExcelInstance.Application.Sheets.Add();
            SheetViewModel svm = new SheetViewModel(ws);
            Assert.AreEqual(ws.Name, svm.DisplayString);

            svm.DisplayString = "HelloWorld";
            Assert.AreEqual(ws.Name, svm.DisplayString,
                "DisplayString is not written through to sheet object.");
        }

        [Test]
        public static void TestWorkbookViewModel()
        {
            Workbook wb = ExcelInstance.Application.Workbooks.Add();
            WorkbookViewModel wvm = new WorkbookViewModel(wb);
            Assert.AreEqual(wvm.DisplayString, wb.Name,
                "WorkbookViewModel does not give workbook name as display string");

            // When accessing sheets in a collection, keep in mind that
            // the Sheets collection of a Workbook instance is 1-based.
            Assert.AreEqual(wvm.Sheets[0].DisplayString, wb.Sheets[1].Name,
                "SheetViewModel in WorkbookViewModel has incorrect sheet name");
        }
    }
}

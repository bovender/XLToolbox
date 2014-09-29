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
    class SheetViewModelTest
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
        public static void Properties()
        {
            Worksheet ws = ExcelInstance.Application.Sheets.Add();
            SheetViewModel svm = new SheetViewModel(ws);
            Assert.AreEqual(ws.Name, svm.DisplayString);

            svm.DisplayString = "HelloWorld";
            Assert.AreEqual(ws.Name, svm.DisplayString,
                "DisplayString is not written through to sheet object.");
        }
    }
}

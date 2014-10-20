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

        [Test]
        // Ruler:           1         2         3
        //         1---5----0---------0---------01
        [TestCase("valid name", true)]
        [TestCase("valid, name; - special !", true)]
        [TestCase("very long name that is invalid because it exceeds 31 characters", false)]
        [TestCase("invalid: characters *", false)]
        [TestCase("/more [invalid] characters\\", false)]
        [TestCase("", false)]
        public static void ValidSheetNames(string testName, bool isValid)
        {
            string s = isValid ? "" : "not ";
            Assert.AreEqual(
                isValid,
                SheetViewModel.IsValidName(testName)
            );
        }
    }
}

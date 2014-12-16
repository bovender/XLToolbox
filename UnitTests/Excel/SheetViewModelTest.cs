using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Excel.Instance;
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
        public void SetUp()
        {
            ExcelInstance.Start();
        }

        [TearDown]
        public void TearDown()
        {
            ExcelInstance.Shutdown();
        }

        [Test]
        public void Properties()
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
        public void ValidSheetNames(string testName, bool isValid)
        {
            string s = isValid ? "" : "not ";
            Assert.AreEqual(
                isValid,
                SheetViewModel.IsValidName(testName)
            );
        }

        [Test]
        public void CountSelectChartSheet()
        {
            Workbook wb = ExcelInstance.CreateWorkbook();
            Worksheet ws = wb.Worksheets.Add();
            SheetViewModel svm = new SheetViewModel(ws);
            Assert.IsFalse(svm.SelectCharts(),
                "SelectCharts() should return false if worksheet does not contain charts.");
            Assert.IsFalse(svm.SelectShapes(),
                "SelectGraphicObjects() should return false if worksheet does not contain any.");
            Chart chart = wb.Charts.Add();
            svm = new SheetViewModel(chart);
            Assert.AreEqual(1, svm.CountCharts());
            Assert.AreEqual(1, svm.CountShapes());
            Assert.IsTrue(svm.SelectCharts(),
                "SelectCharts() should return true if sheet is a chart.");
            Assert.IsTrue(svm.SelectShapes(),
                "SelectGraphicObjects() should return true if sheet is a chart.");
        }

        [Test]
        public void CountSelectChartObjects()
        {
            Workbook wb = ExcelInstance.CreateWorkbook();
            Worksheet ws = wb.Worksheets.Add();
            ChartObjects cos = ws.ChartObjects();
            cos.Add(10, 10, 200, 100);
            cos.Add(250, 10, 200, 100);
            SheetViewModel svm = new SheetViewModel(ws);
            Assert.AreEqual(2, svm.CountCharts(), "CountCharts()");
            Assert.AreEqual(2, svm.CountShapes(), "CountGraphicObjects()");
            Assert.IsTrue(svm.SelectCharts(),
                "SelectCharts() should return true if the sheet contains an embedded chart.");
            Assert.IsTrue(svm.SelectShapes(),
                "SelectGraphicObjects() should return true if the sheet contains an embedded chart.");
        }

        [Test]
        public void CountSelectShapes()
        {
            Workbook wb = ExcelInstance.CreateWorkbook();
            Worksheet ws = wb.Worksheets.Add();
            Shapes shs = ws.Shapes;
            shs.AddLine(10, 10, 20, 30);
            shs.AddLine(50, 50, 20, 30);
            SheetViewModel svm = new SheetViewModel(ws);
            Assert.AreEqual(0, svm.CountCharts());
            Assert.AreEqual(2, svm.CountShapes());
            Assert.IsFalse(svm.SelectCharts(),
                "SelectCharts() should return false if the sheet contains only shapes.");
            Assert.IsTrue(svm.SelectShapes(),
                "SelectGraphicObjects() should return true if the sheet contains shapes.");
        }

        [Test]
        public void CountSelectShapesAndCharts()
        {
            Workbook wb = ExcelInstance.CreateWorkbook();
            Worksheet ws = wb.Worksheets.Add();
            Shapes shs = ws.Shapes;
            shs.AddLine(10, 10, 20, 30);
            shs.AddLine(50, 50, 20, 30);
            ChartObjects cos = ws.ChartObjects();
            cos.Add(10, 10, 200, 100);
            cos.Add(250, 10, 200, 100);
            SheetViewModel svm = new SheetViewModel(ws);
            Assert.AreEqual(2, svm.CountCharts());
            Assert.AreEqual(4, svm.CountShapes());
            Assert.IsTrue(svm.SelectCharts(),
                "SelectCharts() should return true if the sheet contains charts and shapes.");
            Assert.IsTrue(svm.SelectShapes(),
                "SelectGraphicObjects() should return true if the sheet contains charts and shapes.");
        }
    }
}

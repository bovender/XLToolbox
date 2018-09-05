/* SheetViewModelTest.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using System.Runtime.InteropServices;

namespace XLToolbox.Test.Excel
{
    /// <summary>
    /// Unit tests for the XLToolbox.Core.Excel namespace.
    /// </summary>
    [TestFixture]
    class SheetViewModelTest
    {
        [Test]
        public void DisplayString()
        {
            Sheets sheets = Instance.Default.Application.Sheets;
            Worksheet ws = sheets.Add();
            SheetViewModel svm = new SheetViewModel(ws);
            Assert.AreEqual(ws.Name, svm.DisplayString);

            svm.DisplayString = "HelloWorld";
            Assert.AreEqual(ws.Name, svm.DisplayString,
                "DisplayString is not written through to sheet object.");

            if (Marshal.IsComObject(sheets)) Marshal.ReleaseComObject(sheets);
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
            Workbook wb = Instance.Default.CreateWorkbook();
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
            Workbook wb = Instance.Default.CreateWorkbook();
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
            Workbook wb = Instance.Default.CreateWorkbook();
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
            Workbook wb = Instance.Default.CreateWorkbook();
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

        [Test]
        [TestCase("Sheet1", false)]
        [TestCase("Sheet_1", false)]
        [TestCase("Sheet-1", true)]
        [TestCase("Shöüät1", false)]
        [TestCase("Shéèt1", false)]
        [TestCase("Sheet;1", true)]
        [TestCase("Sheet,1", true)]
        [TestCase("Sheet 1", true)]
        public void RefNeedsQuoting(string sheetName, bool expected)
        {
            Assert.AreEqual(expected, SheetViewModel.RequiresQuote(sheetName), sheetName);
        }
    }
}

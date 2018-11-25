/* SelectionViewModelTest.cs
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
using NUnit.Framework;
using Microsoft.Office.Interop.Excel;
using Mso = Microsoft.Office.Core;
using XLToolbox.Excel.ViewModels;
using System.Windows;

namespace XLToolbox.UnitTests.Excel
{
    [TestFixture]
    class SelectionViewModelTest
    {
        SelectionViewModel svm;
        Worksheet ws;

        [SetUp]
        public void StartExcel()
        {
            svm = new SelectionViewModel(Instance.Default.Application);
            Instance.Default.CreateWorkbook();
            ws = Instance.Default.Application.ActiveWorkbook.Worksheets[1];
        }

        [Test]
        public void RangeSelection()
        {
            ws.Range["B10:D20"].Select();
            Rect r = svm.Bounds;
            Console.WriteLine(r);
            Assert.AreEqual(ws.Columns["B"].Width + ws.Columns["C"].Width + ws.Columns["D"].Width, r.Width);
        }

        [Test]
        public void TwoChartsSelection()
        {
            ChartObjects cos = ws.ChartObjects();
            ChartObject co1 = cos.Add(10, 10, 150, 100);
            ChartObject co2 = cos.Add(150, 150, 100, 60);
            Rect boundingRect = new Rect(10, 10, 240, 200);
            co1.Select(true);
            co2.Select(false);
            Rect r = svm.Bounds;
            Console.WriteLine(r);
            Assert.AreEqual(boundingRect, r,
                "Incorrect bounding rectangle for multiple selected charts.");
        }

        [Test]
        public void TwoShapesSelection()
        {
            Shapes shs = ws.Shapes;
            Shape sh1 = shs.AddLine(100, 100, 300, 200);
            Shape sh2 = shs.AddTextbox(
                Mso.MsoTextOrientation.msoTextOrientationHorizontal,
                200, 200, 500, 400);
            Rect boundingRect = new Rect(100, 100, 600, 500);
            sh1.Select(true);
            sh2.Select(false);
            Rect r = svm.Bounds;
            Console.WriteLine(r);
            Assert.AreEqual(boundingRect, r,
                "Incorrect bounding rectangle for multiple selected shapes.");
        }

        [Test]
        [Apartment(System.Threading.ApartmentState.STA)]
        public void CopyRangeToClipboard()
        {
            Clipboard.SetText("xltoolbox test", TextDataFormat.Text);
            svm.CopyToClipboard();
            // Check if the clipboard contains Excel's Biff12 format
            Assert.IsTrue(Clipboard.ContainsData("Biff12"));
            Clipboard.Clear();
        }

        [Test]
        [Apartment(System.Threading.ApartmentState.STA)]
        public void CopyChartsToClipboard()
        {
            Clipboard.SetText("xltoolbox test", TextDataFormat.Text);
            ChartObjects cos = ws.ChartObjects();
            ChartObject co1 = cos.Add(10, 10, 150, 100);
            ChartObject co2 = cos.Add(150, 150, 100, 60);
            co1.Select(true);
            co2.Select(false);
            svm.CopyToClipboard();
            Assert.IsTrue(Clipboard.ContainsData("EnhancedMetaFile"));
            Clipboard.Clear();
        }

        [Test]
        [Apartment(System.Threading.ApartmentState.STA)]
        public void SaveToEmf()
        {
            svm.SaveToEmf(System.IO.Path.GetTempFileName());
        }
    }
}

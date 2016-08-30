/* RangeViewModelTest.cs
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
using System.IO;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using Bovender.Extensions;
using XLToolbox.Excel.Models;
using XLToolbox.Excel.ViewModels;

namespace XLToolbox.Test.Excel
{
    [TestFixture]
    class ReferenceTest
    {
        [TestFixtureSetUp]
        public void TestFixtureSetup()
        {
            Bovender.Logging.LogFile.Default.EnableDebugLogging();
        }

        [Test]
        [TestCase("A1", "", "", "A1")]
        [TestCase("$A1:D$5", "", "", "$A1:D$5")]
        [TestCase("Sheet3!A2:B4", "", "Sheet3", "A2:B4")]
        [TestCase("[Workbook7]Sheet3!AB12:DC23", "Workbook7", "Sheet3", "AB12:DC23")]
        [TestCase(@"c:\in\another\dir\[Workbook7]Sheet3!AB12:DC23", @"c:\in\another\dir\Workbook7", "Sheet3", "AB12:DC23")]
        [TestCase(@"'c:\in\another\dir\[Workbook-7]Sheet3'!AB12:DC23", @"c:\in\another\dir\Workbook-7", "Sheet3", "AB12:DC23")]
        public void SetReference(string reference, string workbook, string worksheet, string address)
        {
            Reference rvm = new Reference(reference);
            Assert.IsTrue(rvm.IsValid, "Is valid");
            Assert.AreEqual(workbook, rvm.WorkbookPath, "Workbook path");
            Assert.AreEqual(worksheet, rvm.SheetName, "Sheet name");
            Assert.AreEqual(address, rvm.Address, "Address");
        }

        [Test]
        [TestCase("1A")]
        [TestCase("Workbook]Sheet1!A1")]
        public void SetInvalidReference(string invalidReference)
        {
            Reference rvm = new Reference(invalidReference);
            Assert.IsFalse(rvm.IsValid, "Is invalid");
            Assert.AreEqual(String.Empty, rvm.WorkbookPath, "Workbook path");
            Assert.AreEqual(String.Empty, rvm.SheetName, "Sheet name");
            Assert.AreEqual(String.Empty, rvm.Address, "Address");
            Assert.Throws<InvalidOperationException>(() => { Range r = rvm.Range; });
        }

        [Test]
        public void SetRangeActiveSheet()
        {
            Workbook workbook = Instance.Default.ActiveWorkbook;
            Sheets worksheets = workbook.Worksheets;
            Worksheet worksheet = worksheets[1];
            string address = "$D$5:$G$10";
            Reference rvm = new Reference(address);
            Range range = rvm.Range;
            Assert.AreEqual(address, range.Address);
            Assert.AreEqual(worksheet.Name, range.Worksheet.Name);
            Assert.AreEqual(workbook.Name, range.Worksheet.Parent.Name);
            Bovender.ComHelpers.ReleaseComObject(worksheet);
            Bovender.ComHelpers.ReleaseComObject(worksheets);
        }

        [Test]
        public void BuildReferenceString()
        {
            string dir = System.IO.Path.GetTempPath();
            string fn = Path.GetRandomFileName();
            Workbook workbook = Instance.Default.ActiveWorkbook;
            workbook.SaveAs(Path.Combine(dir, fn));
            Worksheet worksheet = workbook.ActiveSheet as Worksheet;
            Range range = worksheet.Range["D5:K10"];
            Reference reference = new Reference(range);
            Assert.AreEqual(
                String.Format("{0}[{1}]{2}!{3}",  dir, fn, worksheet.Name, "$D$5:$K$10"),
                reference.ReferenceString);
            workbook.Close(SaveChanges: false);
            System.IO.File.Delete(Path.Combine(dir, fn));
            Bovender.ComHelpers.ReleaseComObject(worksheet);
        }
    }
}

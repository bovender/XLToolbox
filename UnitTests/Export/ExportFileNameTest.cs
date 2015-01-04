/* ExportFileNameTest.cs
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
using NUnit.Framework;
using XLToolbox.Excel.Instance;
using XLToolbox.Export;
using XLToolbox.Export.Models;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    class ExportFileNameTest
    {
        Workbook wb;
        Worksheet ws;
        
        [SetUp]
        public void SetUp()
        {
            ExcelInstance.Start();
            wb = ExcelInstance.CreateWorkbook(1);
            ws = wb.Worksheets[1];
            ws.Name = "helloworld";
        }

        [TearDown]
        public void TearDown()
        {
            ExcelInstance.Shutdown();
        }

        [Test]
        public void GenerateExportFileName()
        {
            string basicTemplate = "{0}_{1}_#{2}.png";
            string exportTemplate = String.Format(basicTemplate,
                "{" + Strings.Workbook + "}",
                "{" + Strings.Worksheet + "}",
                "{" + Strings.Index + "}");
            ExportFileName efn = new ExportFileName(exportTemplate, FileType.Png);
            string result = efn.GenerateNext(ws);
            string expectedResult = String.Format(basicTemplate, wb.Name, ws.Name, "001");
            Assert.AreEqual(expectedResult, result);
            result = efn.GenerateNext(ws);
            expectedResult = String.Format(basicTemplate, wb.Name, ws.Name, "002");
            Assert.AreEqual(expectedResult, result);
        }

        [Test]
        public void GenerateFromTemplateWithoutIndexWithExt()
        {
            string basicTemplate = "templateWithoutIndex{0}.tif";
            string exportTemplate = String.Format(basicTemplate, "");
            ExportFileName efn = new ExportFileName(exportTemplate, FileType.Tiff);
            string result = efn.GenerateNext(ws);
            string expectedResult = String.Format(basicTemplate, "001");
            Assert.AreEqual(expectedResult, result);
            result = efn.GenerateNext(ws);
            expectedResult = String.Format(basicTemplate, "002");
            Assert.AreEqual(expectedResult, result);
        }

        [Test]
        public void GenerateFromTemplateWithoutIndexNoExt()
        {
            FileType ft = FileType.Png;
            string basicTemplate = "templateWithoutIndex{0}";
            string exportTemplate = String.Format(basicTemplate, "");
            ExportFileName efn = new ExportFileName(exportTemplate, ft);
            string result = efn.GenerateNext(ws);
            string expectedResult = String.Format(basicTemplate, "001");
            Assert.AreEqual(
                expectedResult + ft.ToFileNameExtension(), result);
            result = efn.GenerateNext(ws);
            expectedResult = String.Format(
                basicTemplate + ft.ToFileNameExtension(), "002");
            Assert.AreEqual(expectedResult, result);
        }

        [Test]
        public void AddExtensionIfNoneExists()
        {
            FileType ft = FileType.Tiff;
            string basicTemplate = "{0}_{1}_#{2}";
            string exportTemplate = String.Format(basicTemplate,
                "{" + Strings.Workbook + "}",
                "{" + Strings.Worksheet + "}",
                "{" + Strings.Index + "}");
            ExportFileName efn = new ExportFileName(exportTemplate, ft);
            string result = efn.GenerateNext(ws);
            string expectedResult = String.Format(
                basicTemplate + ft.ToFileNameExtension(), wb.Name, ws.Name, "001");
            Assert.AreEqual(expectedResult, result);
            result = efn.GenerateNext(ws);
            expectedResult = String.Format(
                basicTemplate + ft.ToFileNameExtension(), wb.Name, ws.Name, "002");
            Assert.AreEqual(expectedResult, result);
        }

        [Test]
        public void AddNoExtensionIfOneExists()
        {
            FileType ft = FileType.Tiff;
            string basicTemplate = "{0}_{1}_#{2}.tif";
            string exportTemplate = String.Format(basicTemplate,
                "{" + Strings.Workbook + "}",
                "{" + Strings.Worksheet + "}",
                "{" + Strings.Index + "}");
            ExportFileName efn = new ExportFileName(exportTemplate, ft);
            string result = efn.GenerateNext(ws);
            string expectedResult = String.Format(
                basicTemplate, wb.Name, ws.Name, "001");
            Assert.AreEqual(expectedResult, result);
            result = efn.GenerateNext(ws);
            expectedResult = String.Format(
                basicTemplate, wb.Name, ws.Name, "002");
            Assert.AreEqual(expectedResult, result);
        }
    }
}

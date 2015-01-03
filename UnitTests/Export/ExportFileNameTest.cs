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

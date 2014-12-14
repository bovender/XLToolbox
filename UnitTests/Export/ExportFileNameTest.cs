using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using XLToolbox.Excel.Instance;
using XLToolbox.Export;

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
            string basicTemplate = "{0}_{1}_#{2}";
            string exportTemplate = String.Format(basicTemplate,
                "{" + Strings.Workbook + "}",
                "{" + Strings.Worksheet + "}",
                "{" + Strings.Index + "}");
            ExportFileName efn = new ExportFileName(exportTemplate);
            string result = efn.GenerateNext(wb, ws);
            string expectedResult = String.Format(basicTemplate, wb.Name, ws.Name, "001");
            Assert.AreEqual(expectedResult, result);
            result = efn.GenerateNext(wb, ws);
            expectedResult = String.Format(basicTemplate, wb.Name, ws.Name, "002");
            Assert.AreEqual(expectedResult, result);
        }

        [Test]
        public void GenerateFromTemplateWithoutIndexWithExt()
        {
            string basicTemplate = "templateWithoutIndex{0}.tif";
            string exportTemplate = String.Format(basicTemplate, "");
            ExportFileName efn = new ExportFileName(exportTemplate);
            string result = efn.GenerateNext(wb, ws);
            string expectedResult = String.Format(basicTemplate, "001");
            Assert.AreEqual(expectedResult, result);
            result = efn.GenerateNext(wb, ws);
            expectedResult = String.Format(basicTemplate, "002");
            Assert.AreEqual(expectedResult, result);
        }

        [Test]
        public void GenerateFromTemplateWithoutIndexNoExt()
        {
            string basicTemplate = "templateWithoutIndex{0}";
            string exportTemplate = String.Format(basicTemplate, "");
            ExportFileName efn = new ExportFileName(exportTemplate);
            string result = efn.GenerateNext(wb, ws);
            string expectedResult = String.Format(basicTemplate, "001");
            Assert.AreEqual(expectedResult, result);
            result = efn.GenerateNext(wb, ws);
            expectedResult = String.Format(basicTemplate, "002");
            Assert.AreEqual(expectedResult, result);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NUnit.Framework;
using XLToolbox.Core;
using Microsoft.Office.Interop.Excel;

namespace XLToolbox.Test
{
    [TestFixture]
    class TestExcelInstance
    {
        [Test]
        public void StartAndQuitExcel()
        {
            ExcelInstance.Start();
            ExcelInstance.Shutdown();
        }

        [Test]
        public void QuittingExcelWhileInstanceExistsThrows()
        {
            ExcelInstance.Start();
            ExcelInstance.Application.Visible = true;
            Workbook wb = ExcelInstance.Application.Workbooks[1] as Workbook;
            Worksheet sh = wb.Worksheets[1] as Worksheet;
            sh.Cells[1, 1] = "Hello World";
            using (ExcelInstance i = new ExcelInstance())
            {
                Assert.Throws(typeof(ExcelInstanceException), () =>
                {
                    ExcelInstance.Shutdown();
                });
            }
            ExcelInstance.Shutdown();
        }


    }
}

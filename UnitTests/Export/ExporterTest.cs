using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NUnit.Framework;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Export;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Excel.Instance;
using Bovender.Unmanaged;
using XLToolbox.Export.Models;

namespace XLToolbox.Test.Export
{
    [TestFixture]
    class ExporterTest
    {
        [Test]
        [TestCase(FileType.Emf, 0, ColorSpace.Rgb)]
        [TestCase(FileType.Png, 300, ColorSpace.Rgb)]
        [TestCase(FileType.Tiff, 1200, ColorSpace.Monochrome)]
        [TestCase(FileType.Png, 300, ColorSpace.GrayScale)]
        public void ExportChartObject(FileType fileType, int dpi, ColorSpace colorSpace)
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                // ExcelInstance.Application.Visible = true;
                Workbook wb = ExcelInstance.CreateWorkbook();
                Worksheet ws = wb.Worksheets[1];
                ws.Cells[1, 1] = 1;
                ws.Cells[2, 1] = 2;
                ws.Cells[3, 1] = 3;
                ChartObjects cos = ws.ChartObjects();
                ChartObject co = cos.Add(20, 20, 300, 200);
                SeriesCollection sc = co.Chart.SeriesCollection();
                sc.Add(ws.Range["A1:A3"]);
                co.Chart.ChartArea.Select();
                Preset preset = new Preset(fileType, dpi, colorSpace);
                SingleExportSettings settings = new SingleExportSettings(preset,
                    co.Width, co.Height, true);
                settings.FileName = Path.Combine(
                    Path.GetTempPath(),
                    Path.GetTempFileName() + fileType.ToFileNameExtension()
                    );
                File.Delete(settings.FileName);
                Exporter exporter = new Exporter();
                exporter.ExportSelection(settings);
                Assert.IsTrue(File.Exists(settings.FileName));
            }
        }

        [Test]
        public void ExportChartSheet()
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                Workbook wb = ExcelInstance.CreateWorkbook();
                Chart ch = wb.Charts.Add();
                ((_Chart)ch).Activate();
                Preset preset = new Preset(FileType.Png, 300, ColorSpace.Rgb);
                SingleExportSettings settings = new SingleExportSettings();
                settings.Preset = preset;
                settings.FileName = Path.GetFileNameWithoutExtension(Path.GetTempFileName())
                    + preset.FileType.ToFileNameExtension();
                File.Delete(settings.FileName);
                Exporter exporter = new Exporter();
                exporter.ExportSelectionQuick(settings);
                Assert.IsTrue(File.Exists(settings.FileName), "Output file was not created.");
            }
        }

        [Test]
        [RequiresSTA]
        [TestCase(BatchExportScope.ActiveSheet, BatchExportObjects.Charts, BatchExportLayout.SingleItems, 1)]
        [TestCase(BatchExportScope.ActiveWorkbook, BatchExportObjects.Charts, BatchExportLayout.SingleItems, 7)]
        [TestCase(BatchExportScope.ActiveWorkbook, BatchExportObjects.Charts, BatchExportLayout.SheetLayout, 4)]
        [TestCase(BatchExportScope.ActiveSheet, BatchExportObjects.ChartsAndShapes, BatchExportLayout.SingleItems, 4)]
        [TestCase(BatchExportScope.ActiveWorkbook, BatchExportObjects.ChartsAndShapes, BatchExportLayout.SingleItems, 13)]
        [TestCase(BatchExportScope.ActiveWorkbook, BatchExportObjects.ChartsAndShapes, BatchExportLayout.SheetLayout, 4)]
        public void BatchExport(
            BatchExportScope scope, BatchExportObjects objects, 
            BatchExportLayout layout, int expectedNumberOfFiles)
        {
            using (ExcelInstance excel = new ExcelInstance())
            {
                // ExcelInstance.Application.Visible = true;
                Workbook wb = ExcelInstance.CreateWorkbook(3);
                Helpers.CreateSomeCharts(wb.Worksheets[1], 1);
                Helpers.CreateSomeCharts(wb.Worksheets[2], 2);
                Helpers.CreateSomeCharts(wb.Worksheets[3], 3);
                Helpers.CreateSomeShapes(wb.Worksheets[1], 3);
                Helpers.CreateSomeShapes(wb.Worksheets[2], 2);
                Helpers.CreateSomeShapes(wb.Worksheets[3], 1);
                wb.Charts.Add(After: wb.Sheets[wb.Sheets.Count]);
                wb.Sheets[1].Activate();
                BatchExportSettings settings = new BatchExportSettings();
                settings.Path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                Directory.CreateDirectory(settings.Path);
                settings.FileName = "{workbook}_{worksheet}_{index}";
                settings.Preset = new Preset(FileType.Png, 300, ColorSpace.Rgb);
                settings.Layout = layout;
                settings.Objects = objects;
                settings.Scope = scope;
                Exporter exporter = new Exporter();
                exporter.ExportBatch(settings);
                Assert.AreEqual(expectedNumberOfFiles,
                    Directory.GetFiles(settings.Path).Length);
                Directory.Delete(settings.Path, true);
            }
        }
    }
}

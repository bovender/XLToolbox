using Microsoft.Office.Interop.Excel;
/* BatchExporter.cs
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
using Bovender.Extensions;
using System.Text;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Export.Models;

namespace XLToolbox.Export
{
    public class BatchExporter : Bovender.Mvvm.Models.ProcessModel
    {
        #region Public properties

        public BatchExportSettings Settings { get; set; }

        public int PercentCompleted
        {
            get
            {
                if (_batchFileName != null && _numTotal != 0)
                {
                    return Convert.ToInt32(100d * _batchFileName.Counter / _numTotal);
                }
                else
                {
                    return 0;
                }
            }
        }

        #endregion

        #region Constructor

        public BatchExporter(BatchExportSettings settings)
            : base()
        {
            Settings = settings;
            _exporter = new Exporter(settings.Preset);
        }

        #endregion

        #region Implementation of ProcessModel
        
        public override bool Execute()
        {
            _batchFileName = new ExportFileName(
                Settings.Path,
                Settings.FileName,
                Settings.Preset.FileType);
            bool result = false;
            try
            {
                Instance.Default.DisableScreenUpdating();
                switch (Settings.Scope)
                {
                    case BatchExportScope.ActiveSheet:
                        _numTotal = CountInSheet(Instance.Default.Application.ActiveSheet);
                        ExportSheet(Instance.Default.Application.ActiveSheet);
                        break;
                    case BatchExportScope.ActiveWorkbook:
                        _numTotal = CountInWorkbook(Instance.Default.ActiveWorkbook);
                        ExportWorkbook(Instance.Default.ActiveWorkbook);
                        break;
                    case BatchExportScope.OpenWorkbooks:
                        _numTotal = CountInAllWorkbooks();
                        ExportAllWorkbooks();
                        break;
                    default:
                        throw new NotImplementedException(String.Format(
                            "Batch export not implemented for {0}",
                            Settings.Scope));
                }
                Instance.Default.EnableScreenUpdating();
            }
            catch (Exception e)
            {
                Instance.Default.EnableScreenUpdating();
                throw e;
            }
            return result;
        }
        
        #endregion

        #region Private export methods

        private void ExportAllWorkbooks()
        {
            foreach (Workbook wb in Instance.Default.Application.Workbooks)
            {
                ExportWorkbook(wb);
                if (IsCancellationRequested) break;
            }
        }

        private void ExportWorkbook(Workbook workbook)
        {
            ((_Workbook)workbook).Activate();
            Sheets sheets = workbook.Sheets;
            Worksheet sheet;
            for (int i = 1; i <= sheets.Count; i++)
            {
                sheet = sheets[i];
                ExportSheet(sheet);
                Bovender.ComHelpers.ReleaseComObject(sheet);
                if (IsCancellationRequested) break;
            }
            Bovender.ComHelpers.ReleaseComObject(sheets);
        }

        private void ExportSheet(dynamic sheet)
        {
            sheet.Activate();
            switch (Settings.Layout)
            {
                case BatchExportLayout.SheetLayout:
                    ExportSheetLayout(sheet);
                    break;
                case BatchExportLayout.SingleItems:
                    ExportSheetItems(sheet);
                    break;
                default:
                    throw new NotImplementedException(
                        String.Format("Export of {0} not implemented.", Settings.Layout)
                        );
            }
        }

        private void ExportSheetLayout(dynamic sheet)
        {
            SheetViewModel svm = new SheetViewModel(sheet);
            switch (Settings.Objects)
            {
                case BatchExportObjects.Charts:
                    svm.SelectCharts();
                    break;
                case BatchExportObjects.ChartsAndShapes:
                    svm.SelectShapes();
                    break;
                default:
                    throw new NotImplementedException(Settings.Objects.ToString());
            }
            _exporter.FileName = _batchFileName.GenerateNext(sheet, Instance.Default.Application.Selection);
            _exporter.Execute();
        }

        private void ExportSheetItems(dynamic sheet)
        {
            SheetViewModel svm = new SheetViewModel(sheet);
            if (svm.IsChart)
            {
                svm.SelectCharts();
                ExportSelection(svm.Sheet);
            }
            else
            {
                switch (Settings.Objects)
                {
                    case BatchExportObjects.Charts:
                        ExportSheetChartItems(svm.Worksheet);
                        break;
                    case BatchExportObjects.ChartsAndShapes:
                        ExportSheetAllItems(svm.Worksheet);
                        break;
                    default:
                        throw new NotImplementedException(
                            "Single-item export not implemented for " + Settings.Objects.ToString());
                }
            }
        }

        private void ExportSheetChartItems(Worksheet worksheet)
        {
            // Must use an index-based for loop here.
            // A foreach loop caused lots of 0x800a03ec errors from Excel
            // (for whatever reason).
            ChartObjects cos = worksheet.ChartObjects();
            for (int i = 1; i <= cos.Count; i++)
            {
                dynamic item = cos.Item(i);
                item.Select();
                ExportSelection(worksheet);
                Bovender.ComHelpers.ReleaseComObject(((object)item));
                if (IsCancellationRequested) break;
            }
            Bovender.ComHelpers.ReleaseComObject(cos);
        }

        private void ExportSheetAllItems(Worksheet worksheet)
        {
            Shapes shapes = worksheet.Shapes;
            Shape shape;
            for (int i = 1; i <= shapes.Count; i++)
            {
                shape = shapes.Item(i);
                shape.Select(true);
                ExportSelection(worksheet);
                Bovender.ComHelpers.ReleaseComObject(shape);
                if (IsCancellationRequested) break;
            }
            Bovender.ComHelpers.ReleaseComObject(shapes);
        }

        private void ExportSelection(dynamic sheet)
        {
            _exporter.FileName = _batchFileName.GenerateNext(sheet, Instance.Default.Application.Selection);
            _exporter.Execute();
        }

        #endregion

        #region Private counting methods

        private int CountInAllWorkbooks()
        {
            int n = 0;
            Workbooks workbooks = Instance.Default.Workbooks;
            for (int i = 1; i <= workbooks.Count; i++)
            {
                Workbook workbook = workbooks[i];
                n += CountInWorkbook(workbook);
                Bovender.ComHelpers.ReleaseComObject(workbook);
            }
            return n;
        }

        private int CountInWorkbook(Workbook workbook)
        {
            int n = 0;
            Sheets worksheets = workbook.Worksheets;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                Worksheet worksheet = worksheets[i];
                n += CountInSheet(worksheet);
                Bovender.ComHelpers.ReleaseComObject(worksheet);
            }
            Bovender.ComHelpers.ReleaseComObject(worksheets);
            return n;
        }

        private int CountInSheet(dynamic worksheet)
        {
            switch (Settings.Layout)
            {
                case BatchExportLayout.SheetLayout:
                    return CountInSheetLayout(worksheet);
                case BatchExportLayout.SingleItems:
                    return CountInSheetItems(worksheet);
                default:
                    throw new NotImplementedException(
                        String.Format("Export of {0} not implemented.", Settings.Layout)
                        );
            }
        }

        /// <summary>
        /// Returns 1 if the <paramref name="worksheet"/> contains at least
        /// one chart or drawing object, since all charts/drawing objects will
        /// be exported together into one file.
        /// </summary>
        /// <param name="worksheet">Worksheet to examine.</param>
        /// <returns>1 if sheet contains charts/drawings, 0 if not.</returns>
        private int CountInSheetLayout(dynamic worksheet)
        {
            SheetViewModel svm = new SheetViewModel(worksheet);
            switch (Settings.Objects)
            {
                case BatchExportObjects.Charts:
                    return svm.CountCharts() > 0 ? 1 : 0;
                case BatchExportObjects.ChartsAndShapes:
                    return svm.CountShapes() > 0 ? 1 : 0;
                default:
                    throw new NotImplementedException(String.Format(
                        "Export of {0} not implemented.", Settings.Objects));
            }
        }

        private int CountInSheetItems(dynamic worksheet)
        {
            SheetViewModel svm = new SheetViewModel(worksheet);
            switch (Settings.Objects)
            {
                case BatchExportObjects.Charts:
                    return svm.CountCharts();
                case BatchExportObjects.ChartsAndShapes:
                    return svm.CountShapes();
                default:
                    throw new NotImplementedException(String.Format(
                        "Export of {0} not implemented.", Settings.Objects));
            }
        }

        #endregion

        #region Private fields

        private Exporter _exporter;
        private int _numTotal;
        private ExportFileName _batchFileName;
	 
	    #endregion
    }
}

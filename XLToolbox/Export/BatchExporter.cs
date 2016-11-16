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
                    int percent = Convert.ToInt32(100d * _batchFileName.Counter / _numTotal);
                    Logger.Info("PercentCompleted: {0}", percent);
                    return percent;
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
            Logger.Info("Execute: Start with template '{0}'", Settings.FileName);
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
            }
            catch (System.Runtime.InteropServices.ExternalException exe)
            {
                Logger.Fatal("Execute: Caught an external exception");
                Logger.Fatal("Execute: External error code: {0}", exe.ErrorCode);
                Logger.Fatal(exe);
            }
            catch (Exception e)
            {
                Logger.Fatal("Execute: Caught an exception");
                Logger.Fatal(e);
                throw e;
            }
            finally
            {
                Instance.Default.EnableScreenUpdating();
            }
            return result;
        }
        
        #endregion

        #region Private export methods

        private void ExportAllWorkbooks()
        {
            Logger.Info("ExportAllWorkbooks");
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
            Logger.Info("ExportWorkbook: {0} sheet(s)", sheets.Count);
            dynamic sheet;
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
            Logger.Info("ExportSheet");
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
                    Logger.Fatal("ExportSheet: Layout '{0}' not implemented!", Settings.Layout);
                    throw new NotImplementedException(
                        String.Format("Export of {0} not implemented.", Settings.Layout)
                        );
            }
        }

        private void ExportSheetLayout(dynamic sheet)
        {
            Logger.Info("ExportSheetLayout");
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
                    Logger.Fatal("ExportSheetLayout: Object type '{0}' not implemented!", Settings.Objects);
                    throw new NotImplementedException(Settings.Objects.ToString());
            }
            _exporter.FileName = _batchFileName.GenerateNext(sheet, Instance.Default.Application.Selection);
            _exporter.Execute();
        }

        private void ExportSheetItems(dynamic sheet)
        {
            Logger.Info("ExportSheetItems");
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
                        Logger.Fatal("ExportSheetItems: Object type '{0}' not implemented!", Settings.Objects);
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
            Logger.Info("ExportSheetChartItems: {0} object(s)", cos.Count);
            for (int i = 1; i <= cos.Count; i++)
            {
                Logger.Info("ExportSheetChartItems: [{0}]", i);
                dynamic item = cos.Item(i);
                item.Select();
                ExportSelection(worksheet);
                Bovender.ComHelpers.ReleaseComObject((object)item);
                if (IsCancellationRequested) break;
            }
            Bovender.ComHelpers.ReleaseComObject(cos);
        }

        private void ExportSheetAllItems(Worksheet worksheet)
        {
            Shapes shapes = worksheet.Shapes;
            Logger.Info("ExportSheetAllItems: {0} item(s)", shapes.Count);
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
            Logger.Info("ExportSelection");
            Logger.Info("ExportSelection: Sheet: {0}", sheet.Name);
            dynamic selection = Instance.Default.Application.Selection;
            _exporter.FileName = _batchFileName.GenerateNext(sheet, selection);
            _exporter.Execute();
        }

        #endregion

        #region Private counting methods

        private int CountInAllWorkbooks()
        {
            Logger.Info("CountInAllWorkbooks: Counting...");
            int n = 0;
            Workbooks workbooks = Instance.Default.Workbooks;
            for (int i = 1; i <= workbooks.Count; i++)
            {
                Workbook workbook = workbooks[i];
                n += CountInWorkbook(workbook);
                Bovender.ComHelpers.ReleaseComObject(workbook);
            }
            Logger.Info("CountInAllWorkbooks: ... {0}", n);
            return n;
        }

        private int CountInWorkbook(Workbook workbook)
        {
            Logger.Info("CountInWorkbook: Counting...");
            int n = 0;
            Sheets sheets = workbook.Sheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic sheet = sheets[i];
                n += CountInSheet(sheet);
                Bovender.ComHelpers.ReleaseComObject(sheet);
            }
            Bovender.ComHelpers.ReleaseComObject(sheets);
            Logger.Info("CountInWorkbook: ... {0}", n);
            return n;
        }

        private int CountInSheet(dynamic sheet)
        {
            Logger.Info("CountInSheet");
            switch (Settings.Layout)
            {
                case BatchExportLayout.SheetLayout:
                    return CountInSheetLayout(sheet);
                case BatchExportLayout.SingleItems:
                    return CountInSheetItems(sheet);
                default:
                    Logger.Fatal("CountInSheet: Layout '{0}' not implemented!", Settings.Layout);
                    throw new NotImplementedException(
                        String.Format("Export of {0} not implemented.", Settings.Layout)
                        );
            }
        }

        /// <summary>
        /// Returns 1 if the <paramref name="sheet"/> contains at least
        /// one chart or drawing object, since all charts/drawing objects will
        /// be exported together into one file.
        /// </summary>
        /// <param name="sheet">Worksheet to examine.</param>
        /// <returns>1 if sheet contains charts/drawings, 0 if not.</returns>
        private int CountInSheetLayout(dynamic sheet)
        {
            Logger.Info("CountInSheetLayout");
            SheetViewModel svm = new SheetViewModel(sheet);
            switch (Settings.Objects)
            {
                case BatchExportObjects.Charts:
                    return svm.CountCharts() > 0 ? 1 : 0;
                case BatchExportObjects.ChartsAndShapes:
                    return svm.CountShapes() > 0 ? 1 : 0;
                default:
                    Logger.Fatal("CountInSheetLayout: Object type '{0}' not implemented!", Settings.Objects);
                    throw new NotImplementedException(String.Format(
                        "Export of {0} not implemented.", Settings.Objects));
            }
        }

        private int CountInSheetItems(dynamic worksheet)
        {
            Logger.Info("CountInSheetItems");
            SheetViewModel svm = new SheetViewModel(worksheet);
            switch (Settings.Objects)
            {
                case BatchExportObjects.Charts:
                    return svm.CountCharts();
                case BatchExportObjects.ChartsAndShapes:
                    return svm.CountShapes();
                default:
                    Logger.Fatal("CountInSheetItems: Object type '{0}' not implemented!", Settings.Objects);
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

        #region Class logger

        private static NLog.Logger Logger { get { return _logger.Value; } }

        private static readonly Lazy<NLog.Logger> _logger = new Lazy<NLog.Logger>(() => NLog.LogManager.GetCurrentClassLogger());

        #endregion
    }
}

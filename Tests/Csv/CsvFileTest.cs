/* CsvFileTest.cs
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
using NUnit.Framework;
using Microsoft.Office.Interop.Excel;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Csv;
using System.Threading.Tasks;

namespace XLToolbox.Test.Csv
{
    [TestFixture]
    class CsvFileTest
    {
        [Test]
        public void ExportSimpleCsv()
        {
            Worksheet ws = Instance.Default.ActiveWorkbook.Worksheets.Add();
            ws.Cells[3, 5] = "hello";
            ws.Cells[3, 6] = "13";
            // For testing we 'hide' a pipe symbol in the field.
            ws.Cells[4, 5] = "wor|d";
            ws.Cells[4, 6] = 88.5;
            CsvExporter model = new CsvExporter();
            string fn = System.IO.Path.GetTempFileName();
            model.FileName = fn;
            model.FieldSeparator = "|";
            // Use a funky decimal separator
            model.DecimalSeparator = "~";
            model.Execute();
            string contents = System.IO.File.ReadAllText(fn);
            string expected = String.Format(
                "hello|13{0}\"wor|d\"|88~5{0}",
                Environment.NewLine);
            Assert.AreEqual(expected, contents);
            System.IO.File.Delete(fn);
        }

        [Test]
        public void ExportLargeCsv()
        {
            Worksheet ws = Instance.Default.ActiveWorkbook.Worksheets.Add();
            ws.Cells[1, 1] = "hello";
            ws.Cells[1000, 16384]= "world";
            CsvExporter model = new CsvExporter();
            CsvExportViewModel vm = new CsvExportViewModel(model);
            string fn = System.IO.Path.GetTempFileName();
            vm.FileName = fn;
            bool progressCompletedRaised = false;
            vm.ProcessFinishedMessage.Sent += (sender, args) =>
            {
                progressCompletedRaised = true;
            };
            vm.StartProcess();
            Task t = new Task(() =>
            {
                while (model.IsProcessing) { }
            });
            t.Start();
            t.Wait(15000);
            if (vm.IsProcessing)
            {
                vm.CancelProcess();
                Assert.Inconclusive("CSV export took too long, did not finish.");
                // Do not delete the file, leave it for inspection
            }
            else
            {
                Assert.IsTrue(progressCompletedRaised,
                    "ProgressCompleted event was not raised");
                System.IO.File.Delete(fn);
            }
        }

        /* Performance test commented out because it is not a real test. */
        // [Test]
        public void CsvExportPerformance()
        {
            // 2.29 s with alpha 13's multiple events
            string method = System.Reflection.MethodInfo.GetCurrentMethod().ToString();
            Worksheet ws = Instance.Default.ActiveWorkbook.Worksheets.Add();
            ws.Cells[1, 1] = "hello";
            ws.Cells[200, 5] = "world";
            CsvExporter model = new CsvExporter();
            CsvExportViewModel vm = new CsvExportViewModel(model);
            string fn = System.IO.Path.GetTempFileName();
            model.FileName = fn;
            bool running = true;
            long start = 0;
            vm.ProcessFinishedMessage.Sent += (sender, args) =>
            {
                Console.WriteLine(method + ": *** Export completed ***");
                long stop = DateTime.Now.Ticks;
                Console.WriteLine(
                    String.Format("{0}: export took {1} seconds.",
                    method,
                    Math.Round((double)(stop - start) / TimeSpan.TicksPerSecond, 3)
                    ));
                running = false;
            };
            Task waitTask = new Task(
                () =>
                {
                    Console.WriteLine(method + ": *** Wait task started ***");
                    while (running) ;
                }
            );
            waitTask.Start();
            start = DateTime.Now.Ticks;
            model.Range = ws.UsedRange;
            vm.StartProcess();
            waitTask.Wait(-1);
        }
        
    }
}

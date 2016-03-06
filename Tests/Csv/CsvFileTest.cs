﻿/* CsvFileTest.cs
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
            CsvFile csv = new CsvFile();
            string fn = System.IO.Path.GetTempFileName();
            csv.FileName = fn;
            csv.FieldSeparator = "|";
            // Use a funky decimal separator
            csv.DecimalSeparator = "~";
            csv.Export();
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
            ws.Cells[1000000, 16384]= "world";
            CsvFile csv = new CsvFile();
            string fn = System.IO.Path.GetTempFileName();
            csv.FileName = fn;
            bool progressChangedRaised = false;
            bool progressCompletedRaised = false;
            csv.ExportProgressCompleted += (sender, args) =>
            {
                progressCompletedRaised = true;
            };
            bool wasAborted = false;
            Task checkCancelTask = new Task(() =>
            {
                csv.Export();
                // If the task times out because the Export
                // takes too long, the line below will not
                // be reached, and wasAborted will remain false.
                wasAborted = true;
            });
            checkCancelTask.Start();
            checkCancelTask.Wait(400); // should be cancelled after 300+ ms.
            Assert.IsTrue(progressChangedRaised,
                "ProgressChanged event was not raised");
            Assert.IsTrue(wasAborted,
                "Process was not aborted");
            Assert.IsTrue(progressCompletedRaised,
                "ProgressCompleted event was not raised");
            System.IO.File.Delete(fn);
        }

        /* Performance method commented out because it is not a real test. */
        [Test]
        public void CsvExportPerformance()
        {
            // 2.29 s with alpha 13's multiple events
            string method = System.Reflection.MethodInfo.GetCurrentMethod().ToString();
            Worksheet ws = Instance.Default.ActiveWorkbook.Worksheets.Add();
            ws.Cells[1, 1] = "hello";
            ws.Cells[200, 5] = "world";
            CsvFile csv = new CsvFile();
            string fn = System.IO.Path.GetTempFileName();
            csv.FileName = fn;
            bool running = true;
            long start = 0;
            csv.ExportProgressCompleted += (sender, args) =>
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
            csv.Export(ws.UsedRange);
            waitTask.Wait(-1);
        }
        
    }
}
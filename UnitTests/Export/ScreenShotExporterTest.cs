/* ScreenShotExporterTest.cs
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
using XLToolbox.Test;
using XLToolbox.Export;

namespace XLToolbox.Test.Export
{
    [TestFixture]
    class ScreenShotExporterTest
    {
        [Test]
        [RequiresThread(System.Threading.ApartmentState.STA)]
        public void ExportScreenShot()
        {
            ExcelInstance.Start();
            Worksheet ws = ExcelInstance.Application.Worksheets[1];
            Helpers.CreateSomeCharts(ws, 1);
            ws.ChartObjects(1).Select();
            ScreenshotExporter exporter = new ScreenshotExporter();
            exporter.ExportSelection(System.IO.Path.GetTempFileName());
        }
    }
}

/* SingleExportSettingsViewModelTest.cs
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
using Bovender.Mvvm.Messaging;
using XLToolbox.Export;
using XLToolbox.Excel.Instance;
using XLToolbox.Export.ViewModels;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    class SingleExportSettingsViewModelTest
    {
        SingleExportSettingsViewModel svm;

        [SetUp]
        public void SetUp()
        {
            ExcelInstance.Start();
            Workbook wb = ExcelInstance.CreateWorkbook();
            Worksheet ws = wb.Worksheets[1];
            ChartObjects cos = ws.ChartObjects();
            cos.Add(10, 10, 300, 200).Select();
            svm = new SingleExportSettingsViewModel(new XLToolbox.Export.Models.Preset());
        }

        [TearDown]
        public void TearDown()
        {
            ExcelInstance.Shutdown();
        }

        [Test]
        public void Dimensions()
        {
            Assert.AreEqual(300, svm.Width, "Export settings width is incorrect.");
            Assert.AreEqual(200, svm.Height, "Export settings height is incorrect.");
        }

        [Test]
        public void PreserveAspectWidth()
        {
            svm.PreserveAspect = false;
            svm.Height = 100;
            svm.Width = 200;
            svm.PreserveAspect = true;
            svm.Height = 200;
            Assert.AreEqual(svm.Height * 2, svm.Width,
                "Width was not correctly changed.");
        }

        [Test]
        public void PreserveAspectHeight()
        {
            svm.PreserveAspect = false;
            svm.Height = 100;
            svm.Width = 200;
            svm.PreserveAspect = true;
            svm.Width = 400;
            Assert.AreEqual(svm.Width / 2, svm.Height,
                "Height was not correctly changed.");
        }
        
        [Test]
        public void ChooseFileNameCommand()
        {
            bool fnMsgSent = false;
            svm.ChooseFileNameMessage.Sent +=
                (object sender, MessageArgs<FileNameMessageContent> args) =>
                {
                    fnMsgSent = true;
                };
            svm.ChooseFileNameCommand.Execute(null);
            Assert.IsTrue(fnMsgSent, "ChooseFileNameMessage was not sent.");
        }
    }
}

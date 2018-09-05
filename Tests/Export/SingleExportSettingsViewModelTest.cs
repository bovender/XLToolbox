/* SingleExportSettingsViewModelTest.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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
using XLToolbox.Export.ViewModels;
using XLToolbox.Excel.ViewModels;
using XLToolbox.Export.Models;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    class SingleExportSettingsViewModelTest
    {
        SingleExportSettingsViewModel svm;

        [SetUp]
        public void SetUp()
        {
            Workbook wb = Instance.Default.CreateWorkbook();
            Worksheet ws = wb.Worksheets[1];
            ChartObjects cos = ws.ChartObjects();
            cos.Add(10, 10, 300, 200).Select();
            // Get a preset from the UserSettings to enforce the settings are loaded now.
            Preset preset = UserSettings.UserSettings.Default.ExportPresets.FirstOrDefault();
            if (preset == null)
            {
                preset = PresetsRepository.Default.First;
            }
            SingleExportSettings settings = SingleExportSettings.CreateForSelection(preset);
            svm = new SingleExportSettingsViewModel(settings);
        }

        [Test]
        public void Dimensions()
        {
            Assert.AreEqual(300, svm.Width, "Export settings width is incorrect.");
            Assert.AreEqual(200, svm.Height, "Export settings height is incorrect.");
        }

        [Test]
        public void DimensionsChartSheet()
        {
            Chart c = Instance.Default.Application.ActiveWorkbook.Charts.Add();
            Preset preset = PresetsRepository.Default.First;
            SingleExportSettings settings = SingleExportSettings.CreateForSelection(preset);
            svm = new SingleExportSettingsViewModel(settings);
            // Assert small differences because rounding errors are possible
            Assert.IsTrue(Math.Abs(c.ChartArea.Width - svm.Width) < 0.000001,
                "Export settings width is incorrect.");
            Assert.IsTrue(Math.Abs(c.ChartArea.Height - svm.Height) < 0.000001,
                "Export settings height is incorrect.");
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

        [Test]
        public void ExportCommandDisabledWithoutSelection()
        {
            Instance.Default.Reset(); // reset Excel; no workbooks open

            Preset preset = PresetsRepository.Default.First;
            SingleExportSettings settings = SingleExportSettings.CreateForSelection(preset);
            SingleExportSettingsViewModel svm = new SingleExportSettingsViewModel(settings);

            Assert.IsFalse(svm.ExportCommand.CanExecute(null),
                "Export command should be disabled if there is no selection.");
            Instance.Default.CreateWorkbook();
            settings = SingleExportSettings.CreateForSelection(preset);
            svm = new SingleExportSettingsViewModel(settings);
            Assert.IsTrue(svm.ExportCommand.CanExecute(null),
                "Export command should be enabled if something is selected.");
        }
    }
}

/* SingleExportSettingsViewModelTestNoSel.cs
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
using XLToolbox.Excel.Instance;
using XLToolbox.Export.ViewModels;

namespace XLToolbox.UnitTests.Export
{
    /// <summary>
    /// Dedicated test class that allows testing the SingleExportSettingsViewModel
    /// without selection in Excel; this class does not have a SetUp method that
    /// creates an Excel instance.
    /// </summary>
    [TestFixture]
    class SingleExportSettingsViewModelTestNoSel
    {
        [Test]
        public void ExportCommandDisabledWithoutSelection()
        {
            SingleExportSettingsViewModel svm;
            using (new ExcelInstance())
            {
                svm = new SingleExportSettingsViewModel();
                ExcelInstance.CreateWorkbook();
                PresetViewModel pvm = new PresetViewModel();
                svm.PresetsRepository.Presets.Add(pvm);
                pvm.IsSelected = true;
                Assert.IsTrue(svm.ExportCommand.CanExecute(null),
                    "Export command should be enabled if something is selected.");
            }
            Assert.IsFalse(svm.ExportCommand.CanExecute(null),
                "Export command should be disabled if there is no selection.");
        }

    }
}

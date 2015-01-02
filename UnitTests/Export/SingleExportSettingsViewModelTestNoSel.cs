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

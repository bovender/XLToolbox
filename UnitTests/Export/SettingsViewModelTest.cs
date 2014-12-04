using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLToolbox.Export;
using NUnit.Framework;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    class SettingsViewModelTest
    {
        [Test]
        [TestCase(FileType.Emf, false)]
        [TestCase(FileType.Svg, false)]
        [TestCase(FileType.Png, true)]
        [TestCase(FileType.Tiff, true)]
        public void DpiDisabledForVectors(FileType fileType, bool dpiEnabled)
        {
            Settings s = new Settings() { FileType = fileType };
            SettingsViewModel svm = new SettingsViewModel(s);
            Assert.AreEqual(dpiEnabled, svm.IsDpiEnabled);
        }

        [Test]
        [TestCase(FileType.Emf, false)]
        [TestCase(FileType.Svg, false)]
        [TestCase(FileType.Png, true)]
        [TestCase(FileType.Tiff, true)]
        public void ColorSpaceDisabledForVectors(FileType fileType, bool csEnabled)
        {
            Settings s = new Settings() { FileType = fileType };
            SettingsViewModel svm = new SettingsViewModel(s);
            Assert.AreEqual(csEnabled, svm.IsColorSpaceEnabled);
        }
    
        [Test]
        public void DefaultNameIsUpdatedWhenSettingsChange()
        {
            SettingsViewModel svm = new SettingsViewModel(new Settings());
            svm.FileType = FileType.Emf;
            string originalName = svm.Name;
            svm.FileType = FileType.Png;
            Assert.AreNotEqual(originalName, svm.Name);
        }

        [Test]
        public void NameIsNotUpdatedOnceEdited()
        {
            string testName = "test name";
            SettingsViewModel svm = new SettingsViewModel(new Settings());
            svm.FileType = FileType.Emf;
            // Simulate manually editing the settings name
            svm.Name = testName;
            svm.FileType = FileType.Png;
            Assert.AreEqual(testName, svm.Name);
        }
    
    }
}

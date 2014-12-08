using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLToolbox.Export;
using NUnit.Framework;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    class PresetsViewModelTest
    {
        [Test]
        [TestCase(FileType.Emf, false)]
        [TestCase(FileType.Svg, false)]
        [TestCase(FileType.Png, true)]
        [TestCase(FileType.Tiff, true)]
        public void DpiDisabledForVectors(FileType fileType, bool dpiEnabled)
        {
            Preset s = new Preset() { FileType = fileType };
            PresetsViewModel svm = new PresetsViewModel(s);
            Assert.AreEqual(dpiEnabled, svm.IsDpiEnabled);
        }

        [Test]
        [TestCase(FileType.Emf, false)]
        [TestCase(FileType.Svg, false)]
        [TestCase(FileType.Png, true)]
        [TestCase(FileType.Tiff, true)]
        public void ColorSpaceDisabledForVectors(FileType fileType, bool csEnabled)
        {
            Preset s = new Preset() { FileType = fileType };
            PresetsViewModel svm = new PresetsViewModel(s);
            Assert.AreEqual(csEnabled, svm.IsColorSpaceEnabled);
        }
    
        [Test]
        public void DefaultNameIsUpdatedWhenSettingsChange()
        {
            PresetsViewModel svm = new PresetsViewModel(new Preset());
            svm.FileType = FileType.Emf;
            string originalName = svm.Name;
            svm.FileType = FileType.Png;
            Assert.AreNotEqual(originalName, svm.Name);
        }

        [Test]
        public void NameIsNotUpdatedOnceEdited()
        {
            string testName = "test name";
            PresetsViewModel svm = new PresetsViewModel(new Preset());
            svm.FileType = FileType.Emf;
            // Simulate manually editing the settings name
            svm.Name = testName;
            svm.FileType = FileType.Png;
            Assert.AreEqual(testName, svm.Name);
        }
    
    }
}

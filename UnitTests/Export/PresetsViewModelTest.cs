/* PresetsViewModelTest.cs
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
using NUnit.Framework;
using XLToolbox.Export.Models;
using XLToolbox.Export.ViewModels;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    class PresetsViewModelTest
    {
        [Test]
        [TestCase(FileType.Emf, false)]
        // [TestCase(FileType.Svg, false)]
        [TestCase(FileType.Png, true)]
        [TestCase(FileType.Tiff, true)]
        public void DpiDisabledForVectors(FileType fileType, bool dpiEnabled)
        {
            Preset s = new Preset() { FileType = fileType };
            PresetViewModel svm = new PresetViewModel(s);
            Assert.AreEqual(dpiEnabled, svm.IsDpiEnabled);
        }

        [Test]
        [TestCase(FileType.Emf, false)]
        // [TestCase(FileType.Svg, false)]
        [TestCase(FileType.Png, true)]
        [TestCase(FileType.Tiff, true)]
        public void ColorSpaceDisabledForVectors(FileType fileType, bool csEnabled)
        {
            Preset s = new Preset() { FileType = fileType };
            PresetViewModel svm = new PresetViewModel(s);
            Assert.AreEqual(csEnabled, svm.IsColorSpaceEnabled);
        }
    
        [Test]
        public void DefaultNameIsUpdatedWhenSettingsChange()
        {
            PresetViewModel svm = new PresetViewModel(new Preset());
            svm.FileType.AsEnum = FileType.Emf;
            string originalName = svm.Name;
            svm.FileType.AsEnum = FileType.Png;
            Assert.AreNotEqual(originalName, svm.Name);
        }

        [Test]
        public void NameIsNotUpdatedOnceEdited()
        {
            string testName = "test name";
            PresetViewModel svm = new PresetViewModel(new Preset());
            svm.FileType.AsEnum = FileType.Emf;
            // Simulate manually editing the settings name
            svm.Name = testName;
            svm.FileType.AsEnum = FileType.Png;
            Assert.AreEqual(testName, svm.Name);
        }
    
        [Test]
        public void ColorManagementStates()
        {
            // These assertions will fail if there is not a RGB and a CMYK profile
            // installed on the test system!
            PresetViewModel pvm = new PresetViewModel();
            pvm.ColorSpace.AsEnum = ColorSpace.Rgb;
            Assert.IsFalse(pvm.UseColorProfile,
                "'Use color profile' was checked by default");
            Assert.IsTrue(pvm.IsUseColorProfileEnabled,
                "'Use color profile' not enabled by default");
            pvm.ColorSpace.AsEnum = ColorSpace.Cmyk;
            Assert.IsTrue(pvm.UseColorProfile,
                "'Use color profile' not checked with CMYK");
            Assert.IsFalse(pvm.IsUseColorProfileEnabled,
                "'Use color profile' not disabled with CMYK");
            pvm.ColorSpace.AsEnum = ColorSpace.Rgb;
            Assert.IsTrue(pvm.UseColorProfile,
                "'Use color profile' was checked after switching back to RGB");
            Assert.IsTrue(pvm.IsUseColorProfileEnabled,
                "'Use color profile' not re-enabled after switching back to RGB");
        }
    }
}

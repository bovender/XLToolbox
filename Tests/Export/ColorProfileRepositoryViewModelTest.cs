/* ColorProfilesRepositoryViewModelTest.cs
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
using XLToolbox.Export.Models;
using XLToolbox.Export.ViewModels;

namespace XLToolbox.Test.Export
{
    [TestFixture]
    class ColorProfileRepositoryViewModelTest
    {
        ColorProfileRepositoryViewModel vm;

        [SetUp]
        public void SetUp()
        {
            vm = new ColorProfileRepositoryViewModel();
        }

        [Test]
        public void HasDefaultSelection()
        {
            foreach (ColorSpace cs in Enum.GetValues(typeof(ColorSpace)))
            {
                vm.ColorSpace = cs;
                if (vm.Profiles.Count > 0)
                {
                    Assert.IsNotNull(vm.SelectedProfile,
                        "No selected profile despite available profiles for " + cs.ToString());
                }
                else
                {
                    Assert.IsNull(vm.SelectedProfile,
                        "No color profiles, but profile selected for " + cs.ToString());
                }
            }
        }

        [Test]
        public void SwitchColorSpace()
        {
            // This test will fail if there are no color profiles installed!
            vm.ColorSpace = ColorSpace.Rgb;
            ColorProfileViewModel rgbProfile = vm.SelectedProfile;
            Assert.IsNotNull(rgbProfile,
                "No RGB profile was automatically selected");
            vm.ColorSpace = ColorSpace.Cmyk;
            ColorProfileViewModel cmykProfile = vm.SelectedProfile;
            Assert.IsNotNull(cmykProfile,
                "No CMYK profile was automatically selected");
            Assert.AreNotEqual(rgbProfile, cmykProfile,
                "Same selected profiles for RGB and CMYK.");
            vm.ColorSpace = ColorSpace.Rgb;
            Assert.AreEqual(rgbProfile, vm.SelectedProfile,
                "Different selected profile after switching back to RGB");
        }
    }
}

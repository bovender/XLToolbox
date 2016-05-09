/* PresetsRepositoryTest.cs
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
using XLToolbox.Export;
using NUnit.Framework;
using XLToolbox.Export.Models;

namespace XLToolbox.UnitTests.Export
{
    [TestFixture]
    public class PresetsRepositoryTest
    {
        [Test]
        public void RetrieveExistingPreset()
        {
            Preset p = new Preset();
            PresetsRepository r = PresetsRepository.Default;
            Preset other = new Preset();
            Assert.AreSame(p, r.FindOrAdd(p));
            Assert.AreNotSame(other, r.FindOrAdd(p));
        }

        [Test]
        public void RetrieveUnknownPreset()
        {
            Preset p = new Preset();
            PresetsRepository r = PresetsRepository.Default;
            Assert.AreEqual(0, r.Presets.Count);
            Preset o = r.FindOrAdd(p);
            Assert.AreSame(p, o);
            Assert.AreEqual(1, r.Presets.Count);
        }
    }
}

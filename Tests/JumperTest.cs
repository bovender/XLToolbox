/* JumperTest.cs
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
using XLToolbox;

namespace XLToolbox.Test
{
    [TestFixture]
    class JumperTest
    {
        [Test]
        [TestCase(@"https://www.xltoolbox.net", false, false, true)]
        [TestCase(@"https://5.45.105.43", false, false, true)]
        [TestCase(@"file:///c:/windows/explorer.exe", false, true, false)]
        [TestCase(@"c:\windows\explorer.exe", false, true, false)]
        [TestCase(@"d:\oes\not\exist", false, true, false)]
        public void Parse(string target, bool isReference, bool isFile, bool isWebUrl)
        {
            Jumper j = new Jumper(target);
            Assert.AreEqual(isReference, j.IsReference, "Reference");
            Assert.AreEqual(isFile, j.IsFile, "File");
            Assert.AreEqual(isWebUrl, j.IsWebUrl, "Web URL");
        }
    }
}

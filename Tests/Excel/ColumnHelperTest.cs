/* ColumnHelperTest.cs
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
using XLToolbox.Excel.Models;

namespace XLToolbox.Test.Excel
{
    [TestFixture]
    class ColumnHelperTest
    {
        [Test]
        [TestCase("A", 1, false)]
        [TestCase("$AA", 27, true)]
        [TestCase("XFD", 16384, false)]
        public void ParseReference(string reference, long number, bool isFixed)
        {
            ColumnHelper c = new ColumnHelper(reference);
            Assert.AreEqual(number, c.Number);
            Assert.AreEqual(isFixed, c.IsFixed);
        }

        [Test]
        [TestCase(1, false, "A")]
        [TestCase(28, true, "$AB")]
        [TestCase(16384, false, "XFD")]
        public void BuildReference(long number, bool isFixed, string reference)
        {
            ColumnHelper c = new ColumnHelper(number, isFixed);
            Assert.AreEqual(reference, c.Reference);
        }
    }
}

/* RowHelperTest.cs
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
    class RowHelperTest
    {
        [Test]
        [TestCase("1", 1, false)]
        [TestCase("$27", 27, true)]
        [TestCase("16384", 16384, false)]
        public void ParseReference(string reference, long number, bool isFixed)
        {
            RowHelper r = new RowHelper(reference);
            Assert.AreEqual(number, r.Number);
            Assert.AreEqual(isFixed, r.IsFixed);
        }

        [Test]
        [TestCase(1, false, "1")]
        [TestCase(28, true, "$28")]
        [TestCase(16384, false, "16384")]
        public void BuildReference(long number, bool isFixed, string reference)
        {
            RowHelper r = new RowHelper(number, isFixed);
            Assert.AreEqual(reference, r.Reference);
        }

        [Test]
        public void CompareLower()
        {
            RowHelper lower = new RowHelper("50");
            RowHelper higher = new RowHelper("100");
            Assert.IsTrue(lower < higher);
        }

        [Test]
        public void CompareHigher()
        {
            RowHelper lower = new RowHelper("50");
            RowHelper higher = new RowHelper("100");
            Assert.IsTrue(higher > lower);
        }
    }
}

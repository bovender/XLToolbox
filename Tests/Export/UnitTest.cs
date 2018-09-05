/* UnitTest.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2018 Daniel Kraus
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

namespace XLToolbox.UnitTests.Export
{
    /// <summary>
    /// Unit tests for the Unit class.
    /// </summary>
    [TestFixture]
    class UnitTest
    {
        [Test]
        [TestCase(1.0, Unit.Inch, 25.4, Unit.Millimeter)]
        [TestCase(1.0, Unit.Inch, 72.0, Unit.Point)]
        [TestCase(1.0, Unit.Inch,  1.0, Unit.Inch)]
        [TestCase( 72, Unit.Point,   25.4, Unit.Millimeter)]
        [TestCase(1.0, Unit.Point,    1.0, Unit.Point)]
        [TestCase( 72, Unit.Point,    1.0, Unit.Inch)]
        [TestCase( 1.0, Unit.Millimeter,  1, Unit.Millimeter)]
        [TestCase(25.4, Unit.Millimeter, 72, Unit.Point)]
        [TestCase(25.4, Unit.Millimeter,  1, Unit.Inch)]
        public void ConvertUnit(double fromValue, Unit fromUnit, double expectedValue, Unit toUnit)
        {
            Assert.AreEqual(expectedValue, fromUnit.ConvertTo(fromValue, toUnit));
        }
    }
}

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
            Assert.AreEqual(expectedValue, Math.Round(fromUnit.ConvertTo(fromValue, toUnit)));
        }
    }
}

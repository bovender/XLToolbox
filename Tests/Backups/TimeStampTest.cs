/* TimeStampTest.cs
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
using XLToolbox.Backups;

namespace XLToolbox.Test.Backups
{
    [TestFixture]
    class TimeStampTest
    {
        [Test]
        [TestCase("Auswertung_2015-12-01_22-55-13.xlsx", 2015, 12, 01, 22, 55, 13)]
        [TestCase("c:\\with\\dirs\\Auswertung_2014-11-01_22-55-13.xlsx", 2014, 11, 01, 22, 55, 13)]
        public void ParseFileName(string fileName, int yr, int mo, int dy, int hr, int mi, int se)
        {
            TimeStamp ts = new TimeStamp(fileName);
            Assert.AreEqual(new DateTime(yr, mo, dy, hr, mi, se), ts.DateTime);
        }

        [Test]
        public void FormatTimeStamp()
        {
            TimeStamp ts = new TimeStamp();
            ts.DateTime = new DateTime(2015, 11, 23, 10, 30, 00);
            Assert.AreEqual("_2015-11-23_10-30-00", ts.ToString());
        }

        [Test]
        public void FileNameWithTimeStamp()
        {
            TimeStamp ts = new TimeStamp("Auswertung_2015-12-01_22-55-13.xlsx");
            Assert.IsTrue(ts.HasValue);
        }

        [Test]
        public void FileNameWithoutTimeStamp()
        {
            TimeStamp ts = new TimeStamp("Auswertung.xlsx");
            Assert.IsFalse(ts.HasValue);
        }
    }
}

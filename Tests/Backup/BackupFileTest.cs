/* BackupFileTest.cs
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
using System.IO;
using NUnit.Framework;
using XLToolbox.Backup;

namespace XLToolbox.Test.Backup
{
    [TestFixture]
    class BackupFileTest
    {
        [TestFixtureSetUp]
        public void TestFixtureSetup()
        {
            Bovender.Logging.LogFile.Default.EnableDebugLogging();
        }

        [Test]
        [TestCase("Auswertung_20151201_225513.xlsx", 2015, 12, 01, 22, 55, 13)]
        [TestCase("c:\\with\\dirs\\Auswertung_20141101_225513.xlsx", 2014, 11, 01, 22, 55, 13)]
        public void PathWithTimeStamp(string fileName, int yr, int mo, int dy, int hr, int mi, int se)
        {
            BackupFile bf = new BackupFile(fileName);
            Assert.AreEqual(yr, bf.Year, "Year");
            Assert.AreEqual(mo, bf.Month, "Month");
            Assert.AreEqual(dy, bf.Day, "Day");
        }

        [Test]
        public void BackupOfToday()
        {
            string s = DateTime.Today.ToString(TimeStamp.FormatPattern);
            BackupFile bf = new BackupFile(String.Format("c:\testfile{0}.xlsx", s));
            Assert.IsTrue(bf.IsOfToday);
        }

        [Test]
        public void BackupOfOtherDay()
        {
            DateTime dt = new DateTime(2011, 10, 8, 9, 12, 15);
            string s = dt.ToString(TimeStamp.FormatPattern);
            BackupFile bf = new BackupFile(String.Format("c:\testfile{0}.xlsx", s));
            Assert.IsFalse(bf.IsOfToday);
        }

        [Test]
        public void DeleteFile()
        {
            string fn = Path.GetTempFileName();
            Assert.IsTrue(File.Exists(fn));
            BackupFile bf = new BackupFile(fn);
            Assert.IsTrue(bf.Delete(), "BackupFile.Delete() should return true");
            Assert.IsFalse(File.Exists(fn));
        }

        [Test]
        public void FailToDeleteFile()
        {
            string fn = @"i:\do\not\exist\asdfasdfasdf.asdf";
            Assert.IsFalse(File.Exists(fn), "Dummy file should not exist!");
            BackupFile bf = new BackupFile(fn);
            Assert.IsFalse(bf.Delete(), "BackupFile.Delete() should return false");
        }

        [Test]
        public void CreateBackup()
        {
            string fn = Path.GetTempFileName();
            string backupDir = Path.GetRandomFileName();
            Directory.CreateDirectory(backupDir);
            DateTime dt = new DateTime(2015, 11, 23, 9, 31, 00);
            File.SetLastWriteTime(fn, dt);
            BackupFile bf = BackupFile.CreateBackup(fn, backupDir);
            Assert.IsNotNull(bf);
            Assert.AreEqual(dt, bf.TimeStamp.DateTime);
            Directory.Delete(backupDir, true);
            File.Delete(fn);
        }
    }
}

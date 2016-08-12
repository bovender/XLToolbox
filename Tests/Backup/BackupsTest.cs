/* BackupFilesTest.cs
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
    class BackupsTest
    {
        [TestFixtureSetUp]
        public void TestFixtureSetup()
        {
            Bovender.Logging.LogFile.Default.EnableDebugLogging();
        }

        [Test]
        public void BackupDir()
        {
            string dir = ".backup";
            Backups bf = new Backups(@"c:\test.xlsx", dir);
            Assert.AreEqual(dir, bf.BackupDir);
        }

        [Test]
        public void BackupPath()
        {
            string baseDir = @"c:\my\fancy\folder";
            Backups bf = new Backups(baseDir + @"\test.xlsx", ".backup");
            Assert.AreEqual(baseDir + @"\.backup", bf.BackupPath);
        }

        [Test]
        public void InvalidBackupPath()
        {
            string baseDir = @"c:\my\fancy\folder";
            Assert.Throws<ArgumentException>(() => {
                Backups bf = new Backups(baseDir + @"\test.xlsx", @"c:\.backup");
            });
        }

        [Test]
        public void EnumerateBackups()
        {
            string tmpPath = Path.GetTempPath();
            string backupDir = Path.GetRandomFileName();
            string basename = "myfile";
            string file = basename + ".xlsx";
            string backupsStub = Path.Combine(tmpPath, backupDir, basename);
            Directory.CreateDirectory(Path.Combine(tmpPath, backupDir));
            File.Create(backupsStub + "_20151008_133212.xlsx").Dispose();
            File.Create(backupsStub + "_20141008_133212.xlsx").Dispose();
            File.Create(backupsStub + "_20151123_160000.xlsx").Dispose();
            File.Create(backupsStub + "_notaback_up0000.xlsx").Dispose();
            Backups b = new Backups(Path.Combine(tmpPath, basename) + ".xlsx", backupDir);
            Assert.AreEqual(3, b.Count);
            Directory.Delete(Path.Combine(tmpPath, backupDir), true);
        }

        [Test]
        public void Purge()
        {
            string tmpPath = Path.GetTempPath();
            string backupDir = Path.GetRandomFileName();
            string basename = "myfile";
            string file = basename + ".xlsx";
            string backupsStub = Path.Combine(tmpPath, backupDir, basename);
            DateTime dt = DateTime.Today;
            string today = dt.ToString("yyyyMMdd");
            string yesterday = dt.AddDays(-1).ToString("yyyyMMdd");
            Directory.CreateDirectory(Path.Combine(tmpPath, backupDir));

            List<string> expected = new List<string>();
            TouchFile(String.Format("{0}_{1}_090000.xlsx", backupsStub, yesterday));
            TouchFile(String.Format("{0}_{1}_100000.xlsx", backupsStub, yesterday), expected);
            TouchFile(String.Format("{0}_{1}_080000.xlsx", backupsStub, today), expected);
            TouchFile(String.Format("{0}_{1}_090000.xlsx", backupsStub, today), expected);
            TouchFile(String.Format("{0}_{1}_100000.xlsx", backupsStub, today), expected);
            TouchFile(String.Format("{0}_{1}_110000.xlsx", backupsStub, today), expected);
            TouchFile(String.Format("{0}_{1}_120000.xlsx", backupsStub, today), expected);
            TouchFile(String.Format("{0}_{1}_100000.xlsx", backupsStub, dt.AddDays(-2).ToString("yyyyMMdd")), expected);
            TouchFile(String.Format("{0}_{1}_100000.xlsx", backupsStub, dt.AddDays(-4).ToString("yyyyMMdd")), expected);
            TouchFile(String.Format("{0}_{1}_090000.xlsx", backupsStub, dt.AddDays(-4).ToString("yyyyMMdd")));
            TouchFile(String.Format("{0}_{1}_100000.xlsx", backupsStub, dt.AddDays(-5).ToString("yyyyMMdd")), expected);
            TouchFile(String.Format("{0}_{1}_100000.xlsx", backupsStub, dt.AddDays(-6).ToString("yyyyMMdd")), expected);
            TouchFile(String.Format("{0}_{1}_100000.xlsx", backupsStub, dt.AddDays(-7).ToString("yyyyMMdd")), expected);
            TouchFile(String.Format("{0}_{1}_100000.xlsx", backupsStub, dt.AddDays(-8).ToString("yyyyMMdd")), expected);
            TouchFile(backupsStub + "_20151008_133212.xlsx", expected);
            TouchFile(backupsStub + "_20151007_120000.xlsx");
            TouchFile(backupsStub + "_20141008_133212.xlsx", expected);
            TouchFile(backupsStub + "_20140515_120000.xlsx", expected);
            TouchFile(backupsStub + "_20140102_073212.xlsx", expected);
            TouchFile(backupsStub + "_20151123_160000.xlsx", expected);

            Backups b = new Backups(Path.Combine(tmpPath, basename) + ".xlsx", backupDir);
            b.Purge();

            List<string> actual = b.Files.Select(f => f.Path).ToList();
            expected.Sort();
            actual.Sort();
            Assert.AreEqual(expected, actual);
            Directory.Delete(Path.Combine(tmpPath, backupDir), true);
        }

        private void TouchFile(string fn)
        {
            File.Create(fn).Dispose();
        }

        private void TouchFile(string fn, List<string> rememberList)
        {
            TouchFile(fn);
            rememberList.Add(fn);
        }
    }
}

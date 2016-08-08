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
            File.Create(backupsStub + "_2015-10-08_13-32-12.xlsx").Dispose();
            File.Create(backupsStub + "_2014-10-08_13-32-12.xlsx").Dispose();
            File.Create(backupsStub + "_2015-11-23_16-00-00.xlsx").Dispose();
            File.Create(backupsStub + "_nota-ba-ck_up-00-00.xlsx").Dispose();
            Backups b = new Backups(Path.Combine(tmpPath, basename) + ".xlsx", backupDir);
            Assert.AreEqual(3, b.Count);
            Directory.Delete(Path.Combine(tmpPath, backupDir), true);
        }
    }
}

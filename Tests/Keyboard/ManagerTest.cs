/* ManagerTest.cs
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
using XLToolbox.Keyboard;

namespace XLToolbox.Test.Keyboard
{
    [TestFixture]
    class ManagerTest
    {
        [TearDown]
        public void TearDown()
        {
            XLToolbox.Excel.ViewModels.Instance.Default.Dispose();
        }

        // [Test]
        // public void AddShortcut()
        // {
        //     Manager m = Manager.Default;
        //     string keySequence = "^C";
        //     Command command = Command.SaveCsv;
        //     m.AddShortcut(keySequence, command);
        //     int index = m.Shortcuts.Count - 1;
        //     Assert.AreEqual(command, m.Shortcuts[index].Command);
        //     Assert.AreEqual(keySequence, m.Shortcuts[index].KeySequence);
        // }
        // 
        // [Test]
        // public void ReplaceShortcut()
        // {
        //     Manager m = Manager.Default;
        //     m.AddShortcut("test", Command.OpenFromCell);
        //     int count = m.Shortcuts.Count;
        //     m.AddShortcut("test", Command.PointChart);
        //     Assert.AreEqual(count, m.Shortcuts.Count);
        //     Assert.AreEqual(Command.PointChart, m.Shortcuts[count-1].Command);
        // }
        // 
        // [Test]
        // public void DeleteShortcut()
        // {
        //     Manager m = Manager.Default;
        //     Shortcut s = m.AddShortcut("test", Command.OpenFromCell);
        //     m.RemoveShortcut(s);
        //     Assert.IsFalse(m.Shortcuts.Contains(s));
        // }
    }
}

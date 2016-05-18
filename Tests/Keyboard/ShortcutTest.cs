/* ShortcutTest.cs
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
    class ShortcutTest
    {
        [Test]
        [TestCase("A", "A")]
        [TestCase("{F1}", "{F1}")]
        [TestCase("^A", "CONTROL A")]
        [TestCase("+A", "SHIFT A")]
        [TestCase("%A", "ALT A")]
        [TestCase("^+A", "CONTROL SHIFT A")]
        [TestCase("+^A", "SHIFT CONTROL A")]
        [TestCase("{+}", "+")]
        [TestCase("{^}", "^")]
        [TestCase("{%}", "%")]
        public void KeySequenceToHuman(string sequence, string human)
        {
            Shortcut s = new Shortcut(sequence, 0);
            Assert.AreEqual(human, s.HumanKeySequence);
        }
    }
}

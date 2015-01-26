/* MultilineTest.cs
 * part of Daniel's XL Toolbox NG
 * 
 * Copyright 2014-2015 Daniel Kraus
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
using Bovender.Text;

namespace Bovender.UnitTests.Text
{
    [TestFixture]
    class MultilineTest
    {
        Multiline _multiline;

        [SetUp]
        public void SetUp()
        {
            _multiline = new Multiline();
        }

        [Test]
        public void EmptyText()
        {
            Assert.AreEqual(String.Empty, _multiline.Text);
        }

        [Test]
        public void NormalLines()
        {
            _multiline.AddLine("Hello");
            _multiline.AddLine("World");
            Assert.AreEqual("Hello World", _multiline.Text);
        }

        [Test]
        public void LinesWithSpaceAtEnd()
        {
            _multiline.AddLine("Hello          ");
            _multiline.AddLine("World");
            Assert.AreEqual("Hello World", _multiline.Text);
        }

        [Test]
        public void LinesWithLineBreaksAtEnd()
        {
            _multiline.AddLine("Hello\r\n\r\n\r\n");
            _multiline.AddLine("World");
            Assert.AreEqual("Hello World", _multiline.Text);
        }

        [Test]
        public void EmptyLinesInbetween()
        {
            _multiline.AddLine("Hello");
            _multiline.AddLine("");
            _multiline.AddLine("World");
            Assert.AreEqual(3, _multiline.NumberOfLines);
        }

        [Test]
        public void UnorderedListItems()
        {
            _multiline.AddLine("Hello");
            _multiline.AddLine("- This is an unordered list item");
            _multiline.AddLine("     - and this is one with more indent");
            _multiline.AddLine("* List item with asterisk");
            _multiline.AddLine("World");
            Assert.AreEqual(5, _multiline.NumberOfLines);
        }

        [Test]
        public void UnorderedListItemsSeveralLines()
        {
            _multiline.AddLine("Hello");
            _multiline.AddLine("  - This is an unordered list item with indent");
            _multiline.AddLine("    that continues into the next line.");
            _multiline.AddLine("  * And this is another one");
            _multiline.AddLine("    that also continues into the next line.");
            _multiline.AddLine("World");
            Assert.AreEqual(4, _multiline.NumberOfLines);
        }

        [Test]
        public void OrderedListItems()
        {
            _multiline.AddLine("Hello");
            _multiline.AddLine("1. This is an unordered list item");
            _multiline.AddLine("     19) and this is one with more indent");
            _multiline.AddLine("World");
            Assert.AreEqual(4, _multiline.NumberOfLines);
        }

        [Test]
        public void IndentedLines()
        {
            _multiline.AddLine("The first two lines");
            _multiline.AddLine("share the same indent.");
            _multiline.AddLine("  Then there are two lines");
            _multiline.AddLine("  that are indented by two spaces.");
            _multiline.AddLine("And finally a line without indent.");
            Assert.AreEqual(3, _multiline.NumberOfLines);
        }

        [Test]
        public void UnixLineSeparator()
        {
            _multiline.Add("This\nis\na\nUnix\ntext");
            Assert.AreEqual("This is a Unix text", _multiline.Text);
        }

        [Test]
        public void MixedLineSeparators()
        {
            string[] lines = new string[]
            {
                "This test string",             // 0
                "has several lines with",       // 1
                "various different",            // 2
                "line endings, but mostly",     // 3
                "Windows-style line endings",   // 4
                "which consist of a sequence",  // 5
                "of CR+LF."                     // 6
            };
            string expected = String.Join(" ", lines);
            string testText =
                lines[0] + "\r\n" +
                lines[1] + "\n" +
                lines[2] + "\n" +
                lines[3] + "\r\n" +
                lines[4] + "\r\n" +
                lines[5] + "\r" +
                lines[6];
            _multiline.Add(testText);
            Assert.AreEqual(expected, _multiline.Text);
        }

        [Test]
        public void IgnoreComments()
        {
            _multiline.IgnoreHashedComments = true;
            _multiline.AddLine("This line ends # with a comment");
            _multiline.AddLine("# This line should be completeley ignored");
            _multiline.AddLine("beautifully.");
            Assert.AreEqual("This line ends beautifully.", _multiline.Text);
        }
    }
}

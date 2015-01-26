/* Multiline.cs
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
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Bovender.Text
{
    /// <summary>
    /// Represents a text with multiple lines. As lines are added, line breaks
    /// are removed as appropriate.
    /// </summary>
    public class Multiline
    {
        #region Public properties

        /// <summary>
        /// Gets the multiline text as a single string.
        /// </summary>
        public string Text
        {
            get
            {
                if (_dirty)
                {
                    _text = String.Join(Environment.NewLine, _lines);
                    _dirty = false;
                }
                return _text;
            }
        }

        public int NumberOfLines
        {
            get
            {
                return _lines.Count;
            }
        }

        /// <summary>
        /// Indicates whether or not to ignore comments that are
        /// marked with a hash (#).
        /// </summary>
        public bool IgnoreHashedComments { get; set; }

        #endregion

        #region Public methods

        /// <summary>
        /// Adds text to the multiline text. The text to add may
        /// consist of multiple lines. The newline separator to use
        /// will be guessed.
        /// </summary>
        /// <param name="text">Text to add</param>
        public void Add(string text)
        {
            Add(TransformNewLines(text), Environment.NewLine);
        }

        /// <summary>
        /// Adds text to the multiline text by splitting it into
        /// multiple lines by the <paramref name="newLineSeparator"/>.
        /// </summary>
        /// <param name="text">Text to add</param>
        /// <param name="newLineSeparator">Line separator</param>
        public void Add(string text, string newLineSeparator)
        {
            foreach (string line in text.Split(
                new string[] { newLineSeparator }, StringSplitOptions.None))
            {
                AddLine(line);
            }
        }

        /// <summary>
        /// Adds an individual line.
        /// </summary>
        /// <param name="line">Line to add.</param>
        public void AddLine(string line)
        {
            if (IgnoreHashedComments)
            {
                line = Regex.Replace(line, @"#.*$", String.Empty);
                // If the line consisted only of a comment, do not add
                // an empty line.
                if (String.IsNullOrEmpty(line)) return;
            }
            _dirty = true;
            line = line.TrimEnd(' ', '\t', '\n', '\r');
            bool forceNewLine = !HasSameIndent(line);
            forceNewLine |= String.IsNullOrEmpty(line);

            forceNewLine |= IsListItem(line);
            if (forceNewLine || _lines.Count == 0)
            {
                _lines.Add(line);
            }
            else
            {
                _lines[_lines.Count - 1] += " " + line;
            }
        }

        #endregion

        #region Constructors

        public Multiline()
        {
            _dirty = true;
            _lines = new Collection<string>();
        }

        public Multiline(bool ignoreHashedComments)
            : this()
        {
            IgnoreHashedComments = ignoreHashedComments;
        }

        /// <summary>
        /// Creates an instance using <paramref name="text"/> as initial text.
        /// </summary>
        /// <param name="text">Initial text.</param>
        public Multiline(string text)
            : this()
        {
            Add(text);
        }

        public Multiline(string text, bool ignoreHashedComments)
            : this(ignoreHashedComments)
        {
            Add(text);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Determines whether the indent of the current line is the same
        /// as the previous indent and returns true if it is.
        /// </summary>
        /// <param name="line">Line to check</param>
        private bool HasSameIndent(string line)
        {
            Regex r = new Regex(@"^(\s+)");
            Match m = r.Match(line);
            int indent = (m.Success) ? m.Length : 0;
            bool isSame = _lastIndent == indent;
            if (String.IsNullOrEmpty(line))
            {
                // If current line is empty, prevent next line from being
                // appended by setting an impossible indent level.
                _lastIndent = -1;
            }
            else
            {
                _lastIndent = indent;
            }
            return isSame;
        }

        private bool IsListItem(string line)
        {
            Regex r = new Regex(@"^\s*(([-*])|(#+|\d+)\.?\)?)\s");
            Match m = r.Match(line);
            if (m.Success)
            {
                _lastIndent = m.Length;
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Transforms all newlines in a text to the newlines used by
        /// the current environment.
        /// </summary>
        /// <param name="text">Text whose newlines to transform.</param>
        /// <returns>Text with all newlines according to
        /// Environment.NewLine.</returns>
        private string TransformNewLines(string text)
        {
            // Transform everything to "\r\n" first.
            string result = Regex.Replace(text, @"\n\r", "\r\n");
            result = Regex.Replace(result, @"(?<!\r)\n", "\r\n");
            result = Regex.Replace(result, @"\r(?!\n)", "\r\n");

            // Transform the "\r\n" sequences to the one used by the
            // current environment, but only if it is different.
            if (Environment.NewLine != "\r\n")
            {
                result = Regex.Replace(result, @"\r\n", Environment.NewLine);
            }

            return result;
        }

        #endregion

        #region Private fields

        string _text;
        bool _dirty;
        int _lastIndent;
        Collection<String> _lines;

        #endregion
    }
}

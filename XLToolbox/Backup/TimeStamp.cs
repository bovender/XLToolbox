/* TimeStamp.cs
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
using System.Text.RegularExpressions;

namespace XLToolbox.Backup
{
    /// <summary>
    /// Defines a time stamp for use in file names.
    /// </summary>
    public class TimeStamp
    {
        #region Static properties

        /// <summary>
        /// Gets a wildcard pattern for use in file system searches.
        /// </summary>
        public static string WildcardPattern
        {
            get
            {
                // Make sure this always matches the formatting pattern
                // returned by FormatPattern.
                return "_????????_??????";
            }
        }

        /// <summary>
        /// Gets a format string for use in DateTime.ToString().
        /// </summary>
        public static string FormatPattern
        {
            get
            {
                // Make sure this always matches the wildcard pattern
                // returned by WildcardPattern.
                return "_yyyyMMdd_HHmmss";
            }
        }

        /// <summary>
        /// Gets a regular expression that matches a time stamp
        /// in a file name.
        /// </summary>
        public static Regex TimeStampRegex
        {
            get
            {
                if (_regEx == null)
                {
                    // This will only work in the 2nd and 3rd milleniums.
                    // Note for people in the 2990s: Y3K problem! ;-)
                    _regEx = new Regex(@"_[12]\d{7}_\d{6}");
                }
                return _regEx;
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets or sets the date and time represented by this time stamp.
        /// </summary>
        public DateTime DateTime { get; set; }

        public bool HasValue { get; set; }

        #endregion

        #region Constructor

        /// <summary>
        /// Creates a TimeStamp object without information.
        /// </summary>
        public TimeStamp() { }

        /// <summary>
        /// Creates a TimeStamp object by parsing a file name.
        /// </summary>
        /// <param name="fileName"></param>
        public TimeStamp(string fileName)
        {
            Parse(fileName);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Parses a file name and extracts the time stamp.
        /// </summary>
        /// <param name="fileName">File name to parse.</param>
        /// <exception cref="System.FormatException">The file name
        /// contains a pattern that resembles a time stamp, but
        /// is not a valid time stamp.</exception>
        private void Parse(string fileName)
        {
            if (!String.IsNullOrEmpty(fileName))
            {
                Match m = TimeStampRegex.Match(fileName);
                if (m.Success)
                {
                    DateTime = DateTime.ParseExact(m.Value, FormatPattern,
                        System.Globalization.CultureInfo.InvariantCulture);
                    HasValue = true;
                }
                else
                {
                    HasValue = false;
                }
            }
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns the formatted time stamp, or an empty string
        /// if the DateTime property is null.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            if (DateTime == null)
            {
                return String.Empty;
            }
            else
            {
                return DateTime.ToString(FormatPattern);
            }
        }

        #endregion

        #region Static field

        private static Regex _regEx;

        #endregion
    }
}

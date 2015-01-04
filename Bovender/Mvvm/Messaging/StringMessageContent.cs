/* StringMessageContent.cs
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
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace Bovender.Mvvm.Messaging
{
    /// <summary>
    /// Encapsulates a string message that is part of the content
    /// that a view model sends to a consumer (view, test, ...) in
    /// a <see cref="Message"/>.
    /// </summary>
    public class StringMessageContent : MessageContent, IDataErrorInfo
    {
        #region Public properties

        public string Value { get; set; }

        /// <summary>
        /// Delegate function that returns error information on the Value field.
        /// </summary>
        public Func<string, string> Validator { get; set; }

        #endregion

        #region Constructors

        public StringMessageContent() : base() { }

        public StringMessageContent(string initialValue)
            : base()
        {
            Value = initialValue;
        }

        #endregion

        #region IDataErrorInfo implementation

        string IDataErrorInfo.Error
        {
            get { return (this as IDataErrorInfo)["Value"]; }
        }

        string IDataErrorInfo.this[string columnName]
        {
            get
            {
                if (columnName == "Value")
                {
                    return Validator(Value);
                }
                else
                {
                    return String.Empty;
                }
            }
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Disables the Confirm command if the Value is invalid.
        /// </summary>
        /// <returns>True if the Value is valid and the dialog can be closed.</returns>
        protected override bool CanConfirm()
        {
            return String.IsNullOrEmpty((this as IDataErrorInfo)["Value"]);
        }

        #endregion
    }
}

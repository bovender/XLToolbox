using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace XLToolbox.Core.Mvvm
{
    /// <summary>
    /// Encapsulates a string message that is part of the content
    /// that a view model sends to a consumer (view, test, ...) in
    /// a <see cref="Message"/>.
    /// </summary>
    public class StringMessageContent : MessageContent, IDataErrorInfo
    {
        public string Value { get; set; }
        public Func<string, string> Validator { get; set; }

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
    }
}

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
        #region Public properties

        public string Value { get; set; }

        /// <summary>
        /// Delegate function that returns error information on the Value field.
        /// </summary>
        public Func<string, string> Validator { get; set; }

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

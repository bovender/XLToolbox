using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XLToolbox.Core.Mvvm
{
    /// <summary>
    /// Encapsulates a string message that is part of the content
    /// that a view model sends to a consumer (view, test, ...) in
    /// a <see cref="Message"/>.
    /// </summary>
    public class StringMessageContent : MessageContent
    {
        public string Value { get; set; }
    }
}

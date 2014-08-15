using System;
using System.Runtime.Serialization;

namespace XLToolbox.Core
{
    [Serializable]
    public class ExcelInstanceException : Exception
    {
        public ExcelInstanceException() { }
        public ExcelInstanceException(string message) : base(message) { }
        public ExcelInstanceException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public ExcelInstanceException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

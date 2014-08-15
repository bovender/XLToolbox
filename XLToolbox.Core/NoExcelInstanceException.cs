using System;
using System.Runtime.Serialization;

namespace XLToolbox.Core
{
    [Serializable]
    public class NoExcelInstanceException : ExcelInstanceException
    {
        public NoExcelInstanceException() { }
        public NoExcelInstanceException(string message) : base(message) { }
        public NoExcelInstanceException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public NoExcelInstanceException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

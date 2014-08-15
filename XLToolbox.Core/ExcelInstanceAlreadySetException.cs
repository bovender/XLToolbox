using System;
using System.Runtime.Serialization;

namespace XLToolbox.Core
{
    [Serializable]
    public class ExcelInstanceAlreadySetException : Exception
    {
        public ExcelInstanceAlreadySetException() { }
        public ExcelInstanceAlreadySetException(string message) : base(message) { }
        public ExcelInstanceAlreadySetException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public ExcelInstanceAlreadySetException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

using System;
using System.Runtime.Serialization;

namespace XLToolbox.Excel.ViewMmodels
{
    [Serializable]
    public class InvalidSheetNameException : Exception
    {
        public InvalidSheetNameException() { }
        public InvalidSheetNameException(string message) : base(message) { }
        public InvalidSheetNameException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public InvalidSheetNameException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

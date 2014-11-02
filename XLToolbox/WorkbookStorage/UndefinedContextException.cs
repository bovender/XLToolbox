using System;
using System.Runtime.Serialization;

namespace XLToolbox.WorkbookStorage
{
    [Serializable]
    class UndefinedContextException : InvalidContextException
    {
        public UndefinedContextException() { }
        public UndefinedContextException(string message) : base(message) { }
        public UndefinedContextException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public UndefinedContextException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

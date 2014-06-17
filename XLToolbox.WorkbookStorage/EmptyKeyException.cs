using System;
using System.Runtime.Serialization;

namespace XLToolbox.WorkbookStorage
{
    [Serializable]
    public class EmptyKeyException : WorkbookStorageException
    {
        public EmptyKeyException() { }
        public EmptyKeyException(string message) : base(message) { }
        public EmptyKeyException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public EmptyKeyException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

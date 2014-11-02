using System;
using System.Runtime.Serialization;

namespace XLToolbox.WorkbookStorage
{
    [Serializable]
    class UnkownKeyException : WorkbookStorageException
    {
        public UnkownKeyException() { }
        public UnkownKeyException(string message) : base(message) { }
        public UnkownKeyException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public UnkownKeyException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

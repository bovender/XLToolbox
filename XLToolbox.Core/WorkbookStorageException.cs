using System;
using System.Runtime.Serialization;

namespace XLToolbox.Core
{
    [Serializable]
    public class WorkbookStorageException : Exception
    {
        public WorkbookStorageException() { }
        public WorkbookStorageException(string message) : base(message) { }
        public WorkbookStorageException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public WorkbookStorageException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

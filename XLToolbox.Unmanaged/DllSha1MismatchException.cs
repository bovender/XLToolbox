using System;
using System.Runtime.Serialization;

namespace XLToolbox.Unmanaged
{
    [Serializable]
    public class DllSha1MismatchException : Exception
    {
        public DllSha1MismatchException() { }
        public DllSha1MismatchException(string message) : base(message) { }
        public DllSha1MismatchException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public DllSha1MismatchException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

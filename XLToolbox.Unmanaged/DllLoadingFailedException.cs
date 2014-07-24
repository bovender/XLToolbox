using System;
using System.Runtime.Serialization;

namespace XLToolbox.Unmanaged
{
    [Serializable]
    public class DllLoadingFailedException : Exception
    {
        public DllLoadingFailedException() { }
        public DllLoadingFailedException(string message) : base(message) { }
        public DllLoadingFailedException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public DllLoadingFailedException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

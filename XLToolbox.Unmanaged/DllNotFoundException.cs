using System;
using System.Runtime.Serialization;

namespace XLToolbox.Unmanaged
{
    [Serializable]
    public class DllNotFoundException : Exception
    {
        public DllNotFoundException() { }
        public DllNotFoundException(string message) : base(message) { }
        public DllNotFoundException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public DllNotFoundException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

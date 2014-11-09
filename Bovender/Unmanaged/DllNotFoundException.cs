using System;
using System.Runtime.Serialization;

namespace Bovender.Unmanaged
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

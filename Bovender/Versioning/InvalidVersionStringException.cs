using System;
using System.Runtime.Serialization;

namespace Bovender.Versioning
{
    [Serializable]
    public class InvalidVersionStringException : Exception
    {
        public InvalidVersionStringException() { }
        public InvalidVersionStringException(string message) : base(message) { }
        public InvalidVersionStringException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public InvalidVersionStringException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

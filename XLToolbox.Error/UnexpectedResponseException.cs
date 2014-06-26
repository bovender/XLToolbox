using System;
using System.Runtime.Serialization;

namespace XLToolbox.Error
{
    [Serializable]
    public class UnexpectedResponseException : Exception
    {
        public UnexpectedResponseException() { }
        public UnexpectedResponseException(string message) : base(message) { }
        public UnexpectedResponseException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public UnexpectedResponseException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

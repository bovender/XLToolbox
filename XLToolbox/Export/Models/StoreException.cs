using System;
using System.Runtime.Serialization;

namespace XLToolbox.Export.Models
{
    [Serializable]
    public class StoreException : Exception
    {
        public StoreException() { }
        public StoreException(string message) : base(message) { }
        public StoreException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public StoreException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

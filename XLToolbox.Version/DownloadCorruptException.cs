using System;
using System.Runtime.Serialization;

namespace XLToolbox.Version
{
    [Serializable]
    public class DownloadCorruptException : Exception
    {
        public DownloadCorruptException() { }
        public DownloadCorruptException(string message) : base(message) { }
        public DownloadCorruptException(string message,
            Exception innerException)
            : base(message, innerException) { }
        public DownloadCorruptException(SerializationInfo info,
            StreamingContext context)
            : base(info, context) { }
    }
}

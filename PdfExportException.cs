using System;
using System.Runtime.Serialization;

namespace XlsToPDFExporter
{
    [Serializable]
    internal class PdfExportException : Exception
    {
        public PdfExportException()
        {
        }

        public PdfExportException(string message) : base(message)
        {
        }

        public PdfExportException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected PdfExportException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
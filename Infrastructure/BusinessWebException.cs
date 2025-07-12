using System;
using System.Net;

namespace Pro4Soft.SapB1Integration.Infrastructure
{
    public class BusinessWebException : Exception
    {
        public HttpStatusCode? Code { get; }
        public object Payload { get; }

        public BusinessWebException(HttpStatusCode? code) : this(code, null) { }
        public BusinessWebException(string message = null, object payload = null) : this(HttpStatusCode.NotAcceptable, message, payload) { }

        public BusinessWebException(HttpStatusCode? code, string message = null) : base(message)
        {
            Code = code;
        }

        public BusinessWebException(HttpStatusCode? code, string message, object payload = null) : base(message)
        {
            Payload = payload;
            Code = code;
        }
    }
}
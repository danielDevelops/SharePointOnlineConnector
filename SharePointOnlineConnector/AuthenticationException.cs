using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace SharePointOnlineConnector
{
    public class AuthenticationException : Exception
    {
        public string Error { get; }
        public string CorrelationId { get; }
        public string Description { get; }
        public string TraceId { get; }
        public int[] ErrorCodes { get; }
        public DateTime Timestamp { get; }
        public string UserName { get; }
        public AuthenticationException(string authUsername, string jsonError)
            :base(jsonError)
        {
            UserName = authUsername;
            dynamic resp = JsonConvert.DeserializeObject<ExpandoObject>(jsonError, new ExpandoObjectConverter());
            if (Extensions.DynamicHasProperty(resp, "error"))
                Error = resp.error;
            if (Extensions.DynamicHasProperty(resp, "correlation_id"))
                CorrelationId = resp.correlation_id;
            if (Extensions.DynamicHasProperty(resp, "error_description"))
                Description = resp.error_description;
            if (Extensions.DynamicHasProperty(resp, "trace_id"))
                TraceId = resp.trace_id;
            if (Extensions.DynamicHasProperty(resp, "error_codes"))
                ErrorCodes = ((IEnumerable<object>)resp.error_codes).Select(t => Convert.ToInt32(t)).ToArray();
            if (Extensions.DynamicHasProperty(resp, "timestamp"))
                Timestamp = Convert.ToDateTime(resp.timestamp);
        }
        public AuthenticationException(string authUsername,
            string error, 
            string correlationId, 
            string description, 
            string traceid, 
            int[] errorCodes,
            DateTime timestamp)
            :base(description)
        {
            Error = error;
            CorrelationId = correlationId;
            Description = description;
            TraceId = traceid;
            ErrorCodes = errorCodes;
            Timestamp = timestamp;
            UserName = authUsername;
        }
    }
}

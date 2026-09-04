using System;
using System.Net;

namespace Todo.Core.Interfaces
{
    /// <summary>
    /// Thrown when a Graph request fails and retrying it did not help.
    /// </summary>
    /// <remarks>
    /// BUGFIX_REQUESTFAILED: drp090426 - previously a failed request returned its error body to the caller,
    ///  which then parsed it as if it were data and threw a bare NullReferenceException.
    /// </remarks>
    public class GraphRequestException : Exception
    {
        public string RequestUri { get; }

        public HttpStatusCode StatusCode { get; }

        /// <summary>The (possibly truncated) response body, which usually carries the Graph error code.</summary>
        public string ResponseBody { get; }

        public GraphRequestException(string requestUri, HttpStatusCode statusCode, string responseBody)
            : base($"Graph request failed [{(int)statusCode} {statusCode}]: {requestUri}\n{Truncate(responseBody, 1000)}")
        {
            RequestUri = requestUri;
            StatusCode = statusCode;
            ResponseBody = responseBody;
        }

        // TECH: error bodies are small, but a non-Graph failure (proxy, captive portal) can return a whole html page
        private static string Truncate(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;
            return text.Substring(0, maxLength) + "...";
        }
    }
}

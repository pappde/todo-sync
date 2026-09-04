using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using NuGet.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Todo.MSTTool;
using Todo.MSTTool.Auth;
using Todo.Core.Interfaces;

namespace Todo.MSTTool.Graph
{
    public class GraphClient : IGraphClient
    {
        // BUGFIX_REQUESTFAILED: drp090426 - a long "sync all" both outlives its access token (401) and gets throttled
        //  by Graph (429). Neither used to be handled, so the error body was returned as if it were data.
        private const int MaxAttempts = 5;

        private static readonly TimeSpan BaseRetryDelay = TimeSpan.FromSeconds(5);

        public MSTConfiguration Config { get; }

        public IServiceProvider ServiceProvider { get; }

        public AsyncLazy<HttpClient> HttpClientAsync { get; }

        public GraphClient(IServiceProvider serviceProvider)
        {
            ServiceProvider = serviceProvider;
            Config = ServiceProvider.GetService<MSTConfiguration>();

            HttpClientAsync = new AsyncLazy<HttpClient>(CreateHttpClientAsync);
        }

        protected virtual async Task<HttpClient> CreateHttpClientAsync()
        {
            var clientOptions = new GraphClientOptions();
            var client = Microsoft.Graph.GraphClientFactory.Create(clientOptions);
            var auth = ServiceProvider.GetService<InteractiveAuthenticator>();
            await auth.ConfigureClient(client);
            return client;
        }

        // TODO: to Util class
        public bool IsAbsoluteUrl(string url)
        {
            return Uri.TryCreate(url, UriKind.Absolute, out Uri result);
        }

        // uri: is absolute or relative
        // THROWS: GraphRequestException if the request keeps failing
        public async Task<string> RequestAsync(string uri)
        {
            var client = await HttpClientAsync;
            var auth = ServiceProvider.GetService<InteractiveAuthenticator>();

            // TODO: use helper class/function to ensure is properly formatteds
            string requestUri = IsAbsoluteUrl(uri)
                ? uri
                : Config.BaseUri + uri;
            Debug.WriteLine("Request: {0}:", requestUri);

            bool forceTokenRefresh = false;
            for (int attempt = 1; ; attempt++)
            {
                // TECH: attach the token per request (rather than once per HttpClient) so it can be renewed mid-run
                var accessToken = await auth.GetAccessTokenAsync(forceTokenRefresh);
                forceTokenRefresh = false;

                using (var request = new HttpRequestMessage(HttpMethod.Get, requestUri))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    var response = await client.SendAsync(request);
                    var content = await response.Content.ReadAsStringAsync();
                    if (response.IsSuccessStatusCode)
                        return content;

                    if (attempt >= MaxAttempts || !IsRetryable(response.StatusCode))
                        throw new GraphRequestException(requestUri, response.StatusCode, content);

                    // an expired or revoked token needs a new one, not just another wait
                    if (response.StatusCode == HttpStatusCode.Unauthorized)
                        forceTokenRefresh = true;

                    var delay = GetRetryDelay(response, attempt);
                    Console.WriteLine("Request failed [{0}], retrying in {1:0}s [attempt {2}/{3}]: {4}",
                        (int)response.StatusCode, delay.TotalSeconds, attempt, MaxAttempts, requestUri);
                    await Task.Delay(delay);
                }
            }
        }

        private static bool IsRetryable(HttpStatusCode status)
        {
            switch (status)
            {
                case HttpStatusCode.Unauthorized:           // 401 - token expired part way through a long run
                case HttpStatusCode.RequestTimeout:         // 408
                case HttpStatusCode.TooManyRequests:        // 429 - Graph throttling, expected on large lists
                case HttpStatusCode.InternalServerError:    // 500
                case HttpStatusCode.BadGateway:             // 502
                case HttpStatusCode.ServiceUnavailable:     // 503
                case HttpStatusCode.GatewayTimeout:         // 504
                    return true;
                default:
                    return false;
            }
        }

        // PRE: response is a retryable failure
        private static TimeSpan GetRetryDelay(HttpResponseMessage response, int attempt)
        {
            // CODEP: Graph sends Retry-After on 429/503; honor it rather than guessing
            var retryAfter = response.Headers.RetryAfter;
            if (retryAfter != null)
            {
                if (retryAfter.Delta.HasValue && retryAfter.Delta.Value > TimeSpan.Zero)
                    return retryAfter.Delta.Value;

                if (retryAfter.Date.HasValue)
                {
                    var wait = retryAfter.Date.Value - DateTimeOffset.UtcNow;
                    if (wait > TimeSpan.Zero)
                        return wait;
                }
            }

            // exponential backoff: 5s, 10s, 20s, 40s
            return TimeSpan.FromSeconds(BaseRetryDelay.TotalSeconds * Math.Pow(2, attempt - 1));
        }
    }
}

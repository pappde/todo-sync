// TECH:
// 1) Token Cache: stored in runtime folder as "{app}.dll.msalcache.bin3"

using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Abstractions;
using NuGet.Common;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Net.Http;
using Microsoft.Extensions.DependencyInjection;
using System.Diagnostics;
using System.Linq;
using System.CommandLine.Parsing;
using Tavis.UriTemplates;

namespace Todo.MSTTool.Auth
{
    // TODO: interface
    public class InteractiveAuthenticator
    {
        // TECH: renew a little before the token actually expires, so a request never starts with a nearly-dead one
        private static readonly TimeSpan TokenRefreshBuffer = TimeSpan.FromMinutes(5);

        private readonly IPublicClientApplication _app;

        // CODEP: serializes renewals so we can never race two interactive logins
        private readonly SemaphoreSlim _tokenLock = new SemaphoreSlim(1, 1);

        private AuthenticationResult _authResult;

        public MSTConfiguration Config { get; }

        public InteractiveAuthenticator(IServiceProvider serviceProvider)
        {
            Config = serviceProvider.GetService<MSTConfiguration>();

            _app = PublicClientApplicationBuilder
                .Create(Config.ClientId)
                .WithRedirectUri("http://localhost") // Only loopback redirect uri is supported, see https://aka.ms/msal-net-os-browser for details
                .Build();

            // TECH: use LoginHint or IAccount to access
            TokenCacheHelper.EnableSerialization(_app.UserTokenCache);
        }

        private async Task<AuthenticationResult> LoginInteractiveAsync(IPublicClientApplication app)
        {
            Console.WriteLine("Login Interactive");
            var result = await app.AcquireTokenInteractive(Config.Scopes)
                    .ExecuteAsync();
            return result;
        }

        private async Task<AuthenticationResult> LoginAsync(IPublicClientApplication app, IAccount account)
        {
            Console.WriteLine("Login: {0}", account.Username);

            var result = await app.AcquireTokenSilent(Config.Scopes, account)
                              .ExecuteAsync();

            return result;
        }

        // REF: https://learn.microsoft.com/en-us/azure/active-directory/develop/msal-net-acquire-token-silently
        private async Task<AuthenticationResult> GetAuthenticationResultAsync(IPublicClientApplication app)
        {
            // arbitrarily choose the first account actually signed in the token cache
            var accounts = await app.GetAccountsAsync();
            var account = accounts.FirstOrDefault();

            AuthenticationResult result = null;
            if (account != null)
            {
                try
                {
                    result = await LoginAsync(app, account);
                }
                catch (MsalUiRequiredException ex)
                {
                    // A MsalUiRequiredException happened on AcquireTokenSilent.
                    // This indicates you need to call AcquireTokenInteractive to acquire a token
                    Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                    result = await LoginInteractiveAsync(app);
                }
            }
            else
            {
                result = await LoginInteractiveAsync(app);
            }
                

            return result;
        }

        /// <summary>
        /// Return a currently-valid access token, renewing it when it has expired or is close to expiring.
        /// </summary>
        /// <param name="forceRefresh">true to renew even when the cached token still looks valid (e.g. after a 401)</param>
        /// <remarks>
        /// BUGFIX_TOKENEXPIRY: drp090426 - the token used to be acquired once and pinned to the HttpClient, so any
        ///  export running longer than the ~1hr token lifetime started failing every request part way through.
        /// </remarks>
        public async Task<string> GetAccessTokenAsync(bool forceRefresh = false)
        {
            // fast path: cached token is still good, and we were not told to discard it
            if (!forceRefresh && IsTokenUsable(_authResult))
                return _authResult.AccessToken;

            await _tokenLock.WaitAsync();
            try
            {
                // another caller may have renewed it while we waited on the lock
                if (!forceRefresh && IsTokenUsable(_authResult))
                    return _authResult.AccessToken;

                // TECH: AcquireTokenSilent() serves from the MSAL cache and falls back to the refresh token, so this
                //  is cheap and only prompts interactively once the refresh token itself is gone.
                _authResult = await GetAuthenticationResultAsync(_app);
                return _authResult.AccessToken;
            }
            finally
            {
                _tokenLock.Release();
            }
        }

        private static bool IsTokenUsable(AuthenticationResult result)
        {
            return result != null
                && result.ExpiresOn - TokenRefreshBuffer > DateTimeOffset.UtcNow;
        }

        /// <summary>
        /// Prime authentication so any interactive login happens up front rather than part way through a long export.
        /// </summary>
        /// <remarks>CODEP: GraphClient.RequestAsync() attaches a freshly-checked token per request; this is only the initial state.</remarks>
        public async Task ConfigureClient(HttpClient client)
        {
            var accessToken = await GetAccessTokenAsync();
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
        }
    }
}

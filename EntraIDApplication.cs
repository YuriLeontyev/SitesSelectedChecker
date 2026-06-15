using Microsoft.Identity.Client;

namespace SitesSelectedChecker
{
    internal class EntraIDApplication
    {
        private static string[] CsomScopes(Uri tenantRoot) =>
            [
                //$"{tenantRoot.Scheme}://{tenantRoot.Host}/AllSites.Read", // AllSites.Read requires additional permission granted by tenant admin
                $"{tenantRoot.Scheme}://{tenantRoot.Host}/Sites.Selected", // Sites.Selected - works
                //$"{tenantRoot.Scheme}://{tenantRoot.Host}/.default" // Sites.Selected via .default - works
            ];

        private static readonly string[] GraphScopes = ["https://graph.microsoft.com/.default"];

        public EntraIDApplication(Settings settings)
        {
            var builder = ConfidentialClientApplicationBuilder.Create(settings.EntraID.AppId)
                .WithClientSecret(settings.EntraID.AppSecret)
                .WithAuthority($"{settings.EntraID.AuthorityUrl}/{settings.EntraID.TenantId}");

            _clientApp = builder.Build();

            var publlicBuilder = PublicClientApplicationBuilder.Create(settings.EntraID.AppId)
               .WithAuthority($"{settings.EntraID.AuthorityUrl}/{settings.EntraID.TenantId}")
               .WithDefaultRedirectUri();

            _publicClient = publlicBuilder.Build();
        }

        private readonly IConfidentialClientApplication _clientApp;
        private readonly IPublicClientApplication _publicClient;

        public async Task<string?> AuthenticateWithSecretAsync(Settings settings)
        {
            var scopes = new[] { settings.EntraID.ApiScope };

            var loginPrompt = _publicClient
                .AcquireTokenInteractive(scopes)
                .WithPrompt(Prompt.SelectAccount);

            string userToken = (await loginPrompt.ExecuteAsync()).AccessToken;

            AuthenticationResult authResult = await _clientApp
                .AcquireTokenOnBehalfOf(CsomScopes(GetTenantRootUri(settings.SiteUrl)), new UserAssertion(userToken))
                .ExecuteAsync();

            return authResult.AccessToken;
        }

        public async Task<string?> AuthenticateGraphAsync() => await AuthenticatePublicAsync(GraphScopes);

        public async Task<string?> AuthenticateCsomAsync(string siteUrl)
        {
            var tenantRoot = GetTenantRootUri(siteUrl);

            return await AuthenticatePublicAsync(CsomScopes(tenantRoot));
        }

        private async Task<string?> AuthenticatePublicAsync(IEnumerable<string> scopes)
        {
            var loginPrompt = _publicClient
                .AcquireTokenInteractive(scopes)
                .WithPrompt(Prompt.SelectAccount);

            AuthenticationResult authResult = await loginPrompt.ExecuteAsync();

            return authResult.AccessToken;
        }

        private static Uri GetTenantRootUri(string siteUrl)
        {
            var uri = new Uri(siteUrl);
            return new Uri($"{uri.Scheme}://{uri.Host}");
        }
    }
}
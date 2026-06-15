using Microsoft.SharePoint.Client;

namespace SitesSelectedChecker
{
    internal static class AccessChecker
    {
        public static async Task CheckAccessAsync(Settings settings)
        {
            var accessToken = await new EntraIDApplication(settings).AuthenticateWithSecretAsync(settings);

            if (string.IsNullOrEmpty(accessToken))
            {
                Console.WriteLine("Failed to acquire access token.");
                return;
            }

            await CheckAccessToSiteAsync(settings, accessToken);
        }

        private static async Task CheckAccessToSiteAsync(Settings settings, string accessToken)
        {
            using var context = GetContext(settings.SiteUrl, accessToken);

            context.Load(context.Web, w => w.Title, w => w.Url);
            context.Load(context.Web.Lists, inc => inc.Include(l => l.Id, l => l.Title, l => l.Views, l => l.ContentTypesEnabled, l => l.ContentTypes));
            await context.ExecuteQueryAsync();

            Console.WriteLine($"Web Title: {context.Web.Title}");

            if (context.Web.Lists.Count == 0)
            {
                Console.WriteLine("No lists found or access denied.");
                return;
            }

            foreach (var l in context.Web.Lists)
            {
                Console.WriteLine($"List: {l.Title} (ID: {l.Id})");
            }
        }

        public static async Task CheckCsomAccessAsync(Settings settings)
        {
            var accessToken = await new EntraIDApplication(settings).AuthenticateCsomAsync(settings.SiteUrl);

            if (string.IsNullOrEmpty(accessToken))
            {
                Console.WriteLine("Failed to acquire access token.");
                return;
            }

            await CheckAccessToSiteAsync(settings, accessToken);
        }

        public static async Task CheckGraphAccessAsync(Settings settings)
        {
            var accessToken = await new EntraIDApplication(settings).AuthenticateGraphAsync();

            if (string.IsNullOrEmpty(accessToken))
            {
                Console.WriteLine("Failed to acquire access token.");
                return;
            }

            GraphApiClient graphApiClient = new(settings, accessToken);
            string siteId = await graphApiClient.GetSiteIdAsync(settings.SiteUrl);

            var lists = await graphApiClient.GetListsAsync(siteId);
            
            foreach (var l in lists)
            {
                Console.WriteLine($"List: {l.DisplayName} (ID: {l.Id})");
            }
        }

        public static ClientContext GetContext(string siteUrl, string accessToken)
        {
            var context = new ClientContext(siteUrl);

            context.ExecutingWebRequest += (sender, e) =>
            {
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return context;
        }
    }
}

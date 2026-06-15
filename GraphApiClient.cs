using System.Net.Http.Headers;
using System.Text.Json;

namespace SitesSelectedChecker
{
    class GraphApiClient
    {
        private readonly HttpClient _httpClient;
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            PropertyNameCaseInsensitive = true
        };
        private readonly Settings _settings;

        public GraphApiClient(Settings settings, string accessToken)
        {
            _settings = settings;

            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);
            _httpClient.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
        }

        /// <summary>
        /// Resolves a SharePoint site URL to the Graph siteId
        /// (format: hostname,siteCollectionId,webId).
        /// </summary>
        public async Task<string> GetSiteIdAsync(string siteUrl)
        {
            var uri = new Uri(siteUrl);
            // e.g. https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/mysite
            var url = $"{_settings.GraphApiUrl}/sites/{uri.Host}:{uri.AbsolutePath}";
            var json = await _httpClient.GetStringAsync(url);
            var site = JsonSerializer.Deserialize<GraphSite>(json, JsonOptions);
            return site?.Id ?? string.Empty;
        }

        /// <summary>
        /// Returns all lists for the given siteId using
        /// /sites/{siteId}/lists?$select=displayName,id,webUrl,list
        /// </summary>
        public async Task<List<GraphList>> GetListsAsync(string siteId)
        {
            var url = $"{_settings.GraphApiUrl}/sites/{siteId}/lists?$select=displayName,id,webUrl,list,system";
            return [.. (await GetAllPagesAsync<GraphList>(url)).OrderBy(l => l.DisplayName)];
        }

        /// <summary>
        /// Fetches all pages of a Graph collection endpoint, following @odata.nextLink.
        /// </summary>
        private async Task<List<T>> GetAllPagesAsync<T>(string url)
        {
            var results = new List<T>();

            while (!string.IsNullOrEmpty(url))
            {
                var json = await _httpClient.GetStringAsync(url);
                using var document = JsonDocument.Parse(json);
                var root = document.RootElement;

                if (root.TryGetProperty("value", out var valueElement) && valueElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (var itemElement in valueElement.EnumerateArray())
                    {
                        var item = JsonSerializer.Deserialize<T>(itemElement.GetRawText(), JsonOptions);
                        if (item is not null)
                        {
                            results.Add(item);
                        }
                    }
                }

                if (root.TryGetProperty("@odata.nextLink", out var nextLinkElement) && nextLinkElement.ValueKind == JsonValueKind.String)
                {
                    url = nextLinkElement.GetString();
                }
                else
                {
                    url = null;
                }
            }

            return results;
        }
    }
}

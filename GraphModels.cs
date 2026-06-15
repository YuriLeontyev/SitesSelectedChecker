using System.Text.Json.Serialization;

namespace SitesSelectedChecker
{
    class GraphSite
    {
        [JsonPropertyName("id")]
        public required string Id { get; set; }
    }

    class GraphListFacet
    {
        [JsonPropertyName("contentTypesEnabled")]
        public required bool ContentTypesEnabled { get; set; }
    }

    class GraphList
    {
        [JsonPropertyName("id")]
        public required string Id { get; set; }

        [JsonPropertyName("displayName")]
        public required string DisplayName { get; set; }

        [JsonPropertyName("webUrl")]
        public required string WebUrl { get; set; }

        [JsonPropertyName("list")]
        public required GraphListFacet List { get; set; }
    }
}

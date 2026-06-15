namespace SitesSelectedChecker
{
    internal class Settings
    {
        public string SiteUrl { get; set; } = "";
        public string GraphApiUrl { get; set; } = "";
        public EntraIDSettings EntraID { get; set; } = new EntraIDSettings();
    }

    internal class EntraIDSettings
    {
        public string AppId { get; set; } = "";
        public string AppSecret { get; set; } = "";
        
        public string TenantId { get; set; } = "organizations";
        public string AuthorityUrl { get; set; } = "https://login.microsoftonline.com";
        public string ApiScope { get; set; } = "";
    }
}

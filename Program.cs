using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using SitesSelectedChecker;

var builder = Host.CreateApplicationBuilder(args);

builder.Configuration
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
    .AddUserSecrets<Settings>(optional: true, reloadOnChange: true)
    .AddEnvironmentVariables();

builder.Services.Configure<Settings>(builder.Configuration.GetSection(nameof(Settings)));

using var host = builder.Build();

var settings = host.Services.GetRequiredService<IOptions<Settings>>().Value;

try
{
    Console.WriteLine($"Checking access for site: {settings.SiteUrl}");
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"Using CSOM Confidential");
    Console.ForegroundColor = ConsoleColor.White;
    await AccessChecker.CheckAccessAsync(settings);

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"Using Graph API");
    Console.ForegroundColor = ConsoleColor.White;
    await AccessChecker.CheckGraphAccessAsync(settings);

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"Using CSOM Public");
    Console.ForegroundColor = ConsoleColor.White;
    await AccessChecker.CheckCsomAccessAsync(settings);

    Console.WriteLine("Access check completed.");
    Console.ReadLine();
}
catch (Exception ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"Error: {ex.Message}");
    Console.ReadLine();
}
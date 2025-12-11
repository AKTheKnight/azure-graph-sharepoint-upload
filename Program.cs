using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;

var basePath = Directory.GetCurrentDirectory();
if (!File.Exists(Path.Combine(basePath, "appsettings.json")))
{
    basePath = AppContext.BaseDirectory;
}

var configuration = new ConfigurationBuilder()
    .SetBasePath(basePath)
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddJsonFile("appsettings.Development.json", optional: true, reloadOnChange: true)
    .Build();

var tenantId = RequireSetting(configuration, "Graph:TenantId");
var clientId = RequireSetting(configuration, "Graph:ClientId");
var clientSecret = RequireSetting(configuration, "Graph:ClientSecret");
var siteResourceId = RequireSetting(configuration, "Graph:SiteResourceId");

var filePath = args.Length > 0 ? args[0] : "sample.txt";
filePath = Path.GetFullPath(filePath);
if (!File.Exists(filePath))
{
    Console.WriteLine($"File not found: {filePath}");
    Console.WriteLine("Usage: dotnet run -- <path-to-file>");
    return;
}

var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
var graphClient = new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });

try
{
    Console.WriteLine($"Authenticated using app client id {clientId}.");

    var site = await graphClient.Sites[siteResourceId].GetAsync(options =>
    {
        options.QueryParameters.Select = new[] { "id", "name", "displayName", "webUrl" };
    });

    Console.WriteLine($"Found site '{site?.DisplayName ?? site?.Name ?? siteResourceId}'.");

    var drive = await graphClient.Sites[siteResourceId].Drive.GetAsync();
    if (drive?.Id is null)
    {
        Console.WriteLine("The site does not expose a default document library.");
        return;
    }

    await using var fileStream = File.OpenRead(filePath);
    var remoteFileName = BuildRemoteFileName(filePath);

    DriveItem? upload = await graphClient
        .Drives[drive.Id]
        .Items["root"]
        .ItemWithPath(remoteFileName)
        .Content
        .PutAsync(fileStream);

    Console.WriteLine($"Upload complete! View it at: {upload?.WebUrl ?? "unknown location"}");
}
catch (ServiceException serviceEx)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"Graph call failed: {serviceEx.Message}");
    Console.WriteLine($"HTTP status: {(int)serviceEx.ResponseStatusCode} {serviceEx.ResponseStatusCode}");
    Console.ResetColor();
}
catch (Exception ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"Unhandled failure: {ex.Message}");
    Console.WriteLine(ex.ToString());
    Console.ResetColor();
}

static string BuildRemoteFileName(string originalPath)
{
    var name = Path.GetFileNameWithoutExtension(originalPath);
    var extension = Path.GetExtension(originalPath);
    return $"{name}-{DateTimeOffset.UtcNow:yyyyMMddHHmmssfff}{extension}";
}

static string RequireSetting(IConfiguration configuration, string key)
{
    var value = configuration[key];
    if (string.IsNullOrWhiteSpace(value))
    {
        throw new InvalidOperationException($"Missing configuration value '{key}'.");
    }

    return value;
}

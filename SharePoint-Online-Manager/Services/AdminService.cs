using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using SharePointOnlineManager.Models;
using System.Linq;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Service for SharePoint Online Admin API operations.
/// Uses the hidden aggregated site collections list in the admin site.
/// </summary>
public class AdminService : IAdminService
{
    private readonly HttpClient _httpClient;
    private readonly string _adminUrl;
    private readonly string _tenantName;
    private bool _disposed;

    // Hidden list in admin site that contains all site collections
    private const string AggregatedSitesListName = "DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS";

    public AdminService(AuthCookies cookies, string tenantName)
    {
        _tenantName = tenantName;
        _adminUrl = $"https://{tenantName}-admin.sharepoint.com";

        var handler = new HttpClientHandler
        {
            CookieContainer = new CookieContainer(),
            UseCookies = true
        };

        // Add cookies for admin domain
        var adminDomain = $"{tenantName}-admin.sharepoint.com";
        handler.CookieContainer.Add(new Uri(_adminUrl), new Cookie("FedAuth", cookies.FedAuth, "/", adminDomain));
        handler.CookieContainer.Add(new Uri(_adminUrl), new Cookie("rtFa", cookies.RtFa, "/", adminDomain));

        _httpClient = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromMinutes(2)
        };
        _httpClient.DefaultRequestHeaders.Accept.Clear();
        _httpClient.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json")
            {
                Parameters = { new NameValueHeaderValue("odata", "nometadata") }
            });
    }

    public async Task<List<SiteCollection>> GetAllSiteCollectionsAsync(IProgress<string>? progress = null)
    {
        progress?.Report("Fetching site collections from admin site...");

        var sites = new List<SiteCollection>();
        int lastId = 0;
        int totalFetched = 0;
        const int batchSize = 5000;

        // Query the hidden aggregated sites list in the admin site
        // This is what PnP PowerShell uses under the hood
        // Use ID-based pagination for reliability (SharePoint list items have auto-incrementing IDs)
        while (true)
        {
            var url = $"{_adminUrl}/_api/web/lists/GetByTitle('{AggregatedSitesListName}')/items" +
                      $"?$top={batchSize}&$orderby=ID asc&$filter=ID gt {lastId}";

            var response = await _httpClient.GetAsync(url);

            if (!response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException(
                    $"Failed to fetch sites from admin list: {response.StatusCode}\n{content}");
            }

            var json = await response.Content.ReadAsStringAsync();
            var (batch, maxId) = ParseAggregatedSitesResponse(json);

            if (batch.Count == 0)
            {
                // No more items
                break;
            }

            sites.AddRange(batch);
            totalFetched += batch.Count;
            lastId = maxId;

            progress?.Report($"Fetched {totalFetched} site collections...");

            // If we got fewer items than the batch size, we've reached the end
            if (batch.Count < batchSize)
            {
                break;
            }
        }

        // Filter out deleted sites (State != 0 means deleted or locked)
        var activeSites = sites
            .Where(s => !string.IsNullOrEmpty(s.Url))
            .OrderBy(s => s.Url)
            .ToList();

        progress?.Report($"Found {activeSites.Count} site collections");
        return activeSites;
    }

    private static (List<SiteCollection> sites, int maxId) ParseAggregatedSitesResponse(string json)
    {
        var sites = new List<SiteCollection>();
        int maxId = 0;

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        // Get items array
        JsonElement items;
        if (root.TryGetProperty("value", out items))
        {
            // OData format
        }
        else if (root.TryGetProperty("d", out var d) && d.TryGetProperty("results", out items))
        {
            // Verbose OData format
        }
        else
        {
            return (sites, maxId);
        }

        foreach (var item in items.EnumerateArray())
        {
            // Track the ID for pagination
            if (item.TryGetProperty("ID", out var idProp) ||
                item.TryGetProperty("Id", out idProp))
            {
                if (idProp.ValueKind == JsonValueKind.Number)
                {
                    var id = idProp.GetInt32();
                    if (id > maxId) maxId = id;
                }
            }

            var site = new SiteCollection
            {
                Title = GetJsonString(item, "Title"),
                Url = GetJsonString(item, "SiteUrl"),
                Template = GetJsonString(item, "TemplateName"),
                Owner = GetJsonString(item, "SiteOwnerEmail"),
                ExternalSharing = GetJsonString(item, "ExternalSharing"),
                FileCount = GetJsonInt(item, "NumOfFiles"),
                HubSiteId = GetJsonGuid(item, "HubSiteId"),
                GroupId = GetJsonGuid(item, "GroupId"),
                SiteId = GetJsonGuid(item, "SiteId"),
                PageViews = GetJsonInt(item, "PageViews"),
                PagesVisited = GetJsonInt(item, "PagesVisited"),
                State = GetJsonInt(item, "State"),
                ChannelType = GetJsonInt(item, "ChannelType"),
                LanguageLcid = GetJsonInt(item, "LCID")
            };

            // Parse last activity date
            if (item.TryGetProperty("LastActivityOn", out var lastActivity) &&
                lastActivity.ValueKind == JsonValueKind.String &&
                DateTime.TryParse(lastActivity.GetString(), out var lastActivityDate))
            {
                site.LastActivityDate = lastActivityDate;
            }

            // Parse time deleted
            if (item.TryGetProperty("TimeDeleted", out var timeDeleted) &&
                timeDeleted.ValueKind == JsonValueKind.String &&
                DateTime.TryParse(timeDeleted.GetString(), out var timeDeletedDate))
            {
                site.TimeDeleted = timeDeletedDate;
            }

            // Parse storage used (can be float like 6144598.0)
            if (item.TryGetProperty("StorageUsed", out var storageUsed) &&
                storageUsed.ValueKind == JsonValueKind.Number)
            {
                site.StorageUsed = (long)storageUsed.GetDouble();
            }

            // Parse created date
            if (item.TryGetProperty("TimeCreated", out var created) &&
                created.ValueKind == JsonValueKind.String &&
                DateTime.TryParse(created.GetString(), out var createdDate))
            {
                site.CreatedDate = createdDate;
            }

            // Only add if we have a valid URL
            if (!string.IsNullOrEmpty(site.Url))
            {
                // Generate title from URL if missing
                if (string.IsNullOrEmpty(site.Title))
                {
                    site.Title = GetTitleFromUrl(site.Url);
                }

                sites.Add(site);
            }
        }

        return (sites, maxId);
    }

    private static string GetJsonString(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) &&
            prop.ValueKind == JsonValueKind.String)
        {
            return prop.GetString() ?? string.Empty;
        }
        return string.Empty;
    }

    private static int GetJsonInt(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) &&
            prop.ValueKind == JsonValueKind.Number)
        {
            return prop.GetInt32();
        }
        return 0;
    }

    private static Guid GetJsonGuid(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) &&
            prop.ValueKind == JsonValueKind.String &&
            Guid.TryParse(prop.GetString(), out var guid))
        {
            return guid;
        }
        return Guid.Empty;
    }


    private static string GetTitleFromUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length > 0)
            {
                return segments[^1]
                    .Replace("-", " ")
                    .Replace("_", " ");
            }
            return uri.Host;
        }
        catch
        {
            return url;
        }
    }

    public async Task<List<SiteCollection>> GetSiteCollectionsByTypeAsync(SiteType type, IProgress<string>? progress = null)
    {
        var allSites = await GetAllSiteCollectionsAsync(progress);
        return allSites.Where(s => s.SiteType == type).ToList();
    }

    public async Task<bool> TestConnectionAsync()
    {
        try
        {
            // Test by checking if we can access the admin site
            var url = $"{_adminUrl}/_api/web/title";
            var response = await _httpClient.GetAsync(url);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _httpClient.Dispose();
            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }
}

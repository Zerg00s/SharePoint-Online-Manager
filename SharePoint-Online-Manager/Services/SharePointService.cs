using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// SharePoint REST API service using cookie-based authentication.
/// </summary>
public class SharePointService : ISharePointService
{
    private readonly HttpClient _client;
    private readonly string _domain;
    private bool _disposed;

    public string Domain => _domain;

    public SharePointService(AuthCookies cookies) : this(cookies, cookies.Domain)
    {
    }

    /// <summary>
    /// Creates a SharePointService with cookies set for a specific target domain.
    /// This is useful when cookies were obtained from an admin site but need to be used
    /// for regular SharePoint sites in the same tenant.
    /// </summary>
    public SharePointService(AuthCookies cookies, string targetDomain)
    {
        _domain = targetDomain;

        var handler = new HttpClientHandler
        {
            CookieContainer = new CookieContainer()
        };

        // Add cookies for the target domain (FedAuth/rtFa work across the tenant)
        var baseUri = new Uri($"https://{targetDomain}");
        handler.CookieContainer.Add(baseUri, new Cookie("FedAuth", cookies.FedAuth, "/", targetDomain));
        handler.CookieContainer.Add(baseUri, new Cookie("rtFa", cookies.RtFa, "/", targetDomain));

        _client = new HttpClient(handler)
        {
            Timeout = TimeSpan.FromSeconds(60)
        };

        // Use odata=nometadata for simpler JSON structure (no "d" wrapper)
        _client.DefaultRequestHeaders.Accept.Clear();
        _client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json")
            {
                Parameters = { new NameValueHeaderValue("odata", "nometadata") }
            });
    }

    public async Task<SiteInfo> GetSiteInfoAsync(string siteUrl)
    {
        var siteInfo = new SiteInfo { Url = siteUrl };

        try
        {
            var apiUrl = $"{siteUrl.TrimEnd('/')}/_api/web";
            var response = await _client.GetAsync(apiUrl);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            var doc = JsonDocument.Parse(json);

            // Handle both formats: with "d" wrapper (odata=verbose) and without (odata=nometadata)
            var webData = doc.RootElement;
            if (webData.TryGetProperty("d", out var dElement))
            {
                webData = dElement;
            }

            siteInfo.Title = GetStringProperty(webData, "Title");
            siteInfo.Description = GetStringProperty(webData, "Description");
            siteInfo.WebTemplate = GetStringProperty(webData, "WebTemplate");
            siteInfo.Created = GetDateProperty(webData, "Created");
            siteInfo.LastItemModifiedDate = GetDateProperty(webData, "LastItemModifiedDate");
            siteInfo.IsConnected = true;
        }
        catch (HttpRequestException ex)
        {
            siteInfo.IsConnected = false;
            siteInfo.ErrorMessage = ex.StatusCode switch
            {
                HttpStatusCode.Unauthorized => "Authentication failed - cookies may have expired",
                HttpStatusCode.Forbidden => "Access denied - insufficient permissions",
                HttpStatusCode.NotFound => "Site not found",
                _ => $"HTTP error: {ex.Message}"
            };
        }
        catch (Exception ex)
        {
            siteInfo.IsConnected = false;
            siteInfo.ErrorMessage = $"Error: {ex.Message}";
        }

        return siteInfo;
    }

    public async Task<bool> TestConnectionAsync(string siteUrl)
    {
        try
        {
            var apiUrl = $"{siteUrl.TrimEnd('/')}/_api/web/title";
            var response = await _client.GetAsync(apiUrl);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    public Task<SharePointResult<List<ListInfo>>> GetListsAsync(string siteUrl)
    {
        return GetListsAsync(siteUrl, includeHidden: true);
    }

    public async Task<SharePointResult<List<ListInfo>>> GetListsAsync(string siteUrl, bool includeHidden)
    {
        try
        {
            var apiUrl = $"{siteUrl.TrimEnd('/')}/_api/web/lists" +
                         "?$select=Id,Title,ItemCount,Hidden,Created,LastItemModifiedDate,BaseTemplate,RootFolder/ServerRelativeUrl" +
                         "&$expand=RootFolder";

            var response = await _client.GetAsync(apiUrl);

            if (!response.IsSuccessStatusCode)
            {
                return new SharePointResult<List<ListInfo>>
                {
                    Status = response.StatusCode switch
                    {
                        HttpStatusCode.Unauthorized => SharePointResultStatus.AuthenticationRequired,
                        HttpStatusCode.Forbidden => SharePointResultStatus.AccessDenied,
                        HttpStatusCode.NotFound => SharePointResultStatus.NotFound,
                        _ => SharePointResultStatus.Error
                    },
                    ErrorMessage = GetErrorMessage(response.StatusCode)
                };
            }

            var json = await response.Content.ReadAsStringAsync();
            var lists = ParseLists(json);

            if (!includeHidden)
            {
                lists = lists.Where(l => !l.Hidden).ToList();
            }

            return new SharePointResult<List<ListInfo>>
            {
                Data = lists,
                Status = SharePointResultStatus.Success
            };
        }
        catch (HttpRequestException ex)
        {
            return new SharePointResult<List<ListInfo>>
            {
                Status = ex.StatusCode switch
                {
                    HttpStatusCode.Unauthorized => SharePointResultStatus.AuthenticationRequired,
                    HttpStatusCode.Forbidden => SharePointResultStatus.AccessDenied,
                    HttpStatusCode.NotFound => SharePointResultStatus.NotFound,
                    _ => SharePointResultStatus.Error
                },
                ErrorMessage = ex.Message
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<List<ListInfo>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    private static List<ListInfo> ParseLists(string json)
    {
        var lists = new List<ListInfo>();

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            JsonElement valueElement;
            if (root.TryGetProperty("value", out valueElement))
            {
                // Standard OData response
            }
            else if (root.TryGetProperty("d", out var dElement) &&
                     dElement.TryGetProperty("results", out valueElement))
            {
                // OData verbose response
            }
            else
            {
                return lists;
            }

            foreach (var item in valueElement.EnumerateArray())
            {
                var listInfo = new ListInfo
                {
                    Id = GetGuidProperty(item, "Id"),
                    Title = GetStringProperty(item, "Title"),
                    ItemCount = GetIntProperty(item, "ItemCount"),
                    Hidden = GetBoolProperty(item, "Hidden"),
                    Created = GetDateProperty(item, "Created"),
                    LastItemModifiedDate = GetDateProperty(item, "LastItemModifiedDate"),
                    BaseTemplate = GetIntProperty(item, "BaseTemplate")
                };

                // Get RootFolder/ServerRelativeUrl
                if (item.TryGetProperty("RootFolder", out var rootFolder))
                {
                    listInfo.ServerRelativeUrl = GetStringProperty(rootFolder, "ServerRelativeUrl");
                }

                lists.Add(listInfo);
            }
        }
        catch
        {
            // Return what we have
        }

        return lists;
    }

    private static string GetStringProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) && prop.ValueKind == JsonValueKind.String)
            return prop.GetString() ?? string.Empty;
        return string.Empty;
    }

    private static DateTime GetDateProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) && prop.ValueKind == JsonValueKind.String)
        {
            var dateStr = prop.GetString();
            if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out var date))
                return date;
        }
        return DateTime.MinValue;
    }

    private static int GetIntProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) && prop.ValueKind == JsonValueKind.Number)
            return prop.GetInt32();
        return 0;
    }

    private static bool GetBoolProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop))
        {
            if (prop.ValueKind == JsonValueKind.True)
                return true;
            if (prop.ValueKind == JsonValueKind.False)
                return false;
        }
        return false;
    }

    private static Guid GetGuidProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop) && prop.ValueKind == JsonValueKind.String)
        {
            var str = prop.GetString();
            if (!string.IsNullOrEmpty(str) && Guid.TryParse(str, out var guid))
                return guid;
        }
        return Guid.Empty;
    }

    private static string GetErrorMessage(HttpStatusCode statusCode)
    {
        return statusCode switch
        {
            HttpStatusCode.Unauthorized => "Authentication failed - cookies may have expired",
            HttpStatusCode.Forbidden => "Access denied - insufficient permissions",
            HttpStatusCode.NotFound => "Site or resource not found",
            _ => $"HTTP error: {(int)statusCode}"
        };
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _client.Dispose();
            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }
}

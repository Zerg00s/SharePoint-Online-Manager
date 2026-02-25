using System.Collections.Concurrent;
using System.Diagnostics;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Screens;

namespace SharePointOnlineManager.Services;

/// <summary>
/// HTTP handler that logs all request/response status codes to Trace (visible in DebugView).
/// </summary>
internal class TraceLoggingHandler : DelegatingHandler
{
    public TraceLoggingHandler(HttpMessageHandler innerHandler) : base(innerHandler) { }

    protected override async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var method = request.Method;
        var url = request.RequestUri?.AbsoluteUri ?? "unknown";
        // Truncate URL for readability (keep first 120 chars)
        var shortUrl = url.Length > 120 ? url[..120] + "..." : url;

        Trace.WriteLine($"[SPOManager] HTTP {method} {shortUrl}");

        var sw = Stopwatch.StartNew();
        var response = await base.SendAsync(request, cancellationToken);
        sw.Stop();

        var status = (int)response.StatusCode;
        var level = status >= 400 ? "WARN" : "INFO";
        Trace.WriteLine($"[SPOManager] HTTP {status} {response.StatusCode} ({sw.ElapsedMilliseconds}ms) {method} {shortUrl}");

        if (status == 429 || status == 503)
        {
            var retryAfter = response.Headers.RetryAfter?.Delta?.TotalSeconds
                          ?? response.Headers.RetryAfter?.Date?.Subtract(DateTimeOffset.UtcNow).TotalSeconds;
            var retryInfo = retryAfter.HasValue ? $"Retry-After: {retryAfter:F0}s" : "No Retry-After header";
            Trace.WriteLine($"[SPOManager] THROTTLED ({status}) - {retryInfo} - {shortUrl}");
        }
        else if (status >= 400)
        {
            try
            {
                var body = await response.Content.ReadAsStringAsync(cancellationToken);
                if (!string.IsNullOrEmpty(body))
                {
                    var snippet = body.Length > 300 ? body[..300] + "..." : body;
                    Trace.WriteLine($"[SPOManager] HTTP {status} Body: {snippet}");
                }
            }
            catch { /* don't fail on logging */ }
        }

        return response;
    }
}

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

        var cookieHandler = new HttpClientHandler
        {
            CookieContainer = new CookieContainer()
        };

        // Add cookies for the target domain (FedAuth/rtFa work across the tenant)
        var baseUri = new Uri($"https://{targetDomain}");
        cookieHandler.CookieContainer.Add(baseUri, new Cookie("FedAuth", cookies.FedAuth, "/", targetDomain));
        cookieHandler.CookieContainer.Add(baseUri, new Cookie("rtFa", cookies.RtFa, "/", targetDomain));

        // Wrap with trace logging handler for DebugView visibility
        var tracingHandler = new TraceLoggingHandler(cookieHandler);

        _client = new HttpClient(tracingHandler)
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

    public async Task<SharePointResult<List<DocumentReportItem>>> GetDocumentLibraryFilesAsync(
        string siteUrl,
        string libraryTitle,
        bool includeSubfolders = true,
        bool includeVersionCount = true)
    {
        var documents = new List<DocumentReportItem>();
        var baseUrl = siteUrl.TrimEnd('/');
        var encodedLibraryTitle = Uri.EscapeDataString(libraryTitle);
        var errors = new List<string>();

        try
        {
            // Use RenderListDataAsStream for better large library support
            var viewXml = @"<View Scope='RecursiveAll'>
                <Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query>
                <ViewFields>
                    <FieldRef Name='FileLeafRef'/>
                    <FieldRef Name='FileRef'/>
                    <FieldRef Name='File_x0020_Size'/>
                    <FieldRef Name='Created'/>
                    <FieldRef Name='Modified'/>
                    <FieldRef Name='Author'/>
                    <FieldRef Name='Editor'/>
                    <FieldRef Name='_UIVersionString'/>
                </ViewFields>
                <RowLimit Paged='TRUE'>5000</RowLimit>
            </View>";

            string? pagingInfo = null;
            int pageCount = 0;
            var baseApiUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedLibraryTitle}')/RenderListDataAsStream";

            do
            {
                pageCount++;

                // Build the API URL - append paging info if we have it
                var apiUrl = baseApiUrl;
                if (!string.IsNullOrEmpty(pagingInfo))
                {
                    // NextHref contains query params like ?Paged=TRUE&p_ID=5000
                    apiUrl = baseApiUrl + pagingInfo;
                }

                var requestBody = new StringContent(
                    $"{{\"parameters\":{{\"RenderOptions\":2,\"ViewXml\":\"{viewXml.Replace("\"", "\\\"").Replace("\n", "").Replace("\r", "")}\"}}}}",
                    System.Text.Encoding.UTF8,
                    "application/json");

                var response = await _client.PostAsync(apiUrl, requestBody);

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    errors.Add($"Page {pageCount}: HTTP {(int)response.StatusCode} - {errorContent.Substring(0, Math.Min(200, errorContent.Length))}");
                    break;
                }

                var json = await response.Content.ReadAsStringAsync();
                var (items, nextPaging) = ParseRenderListDataResponse(json, siteUrl, libraryTitle);
                documents.AddRange(items);
                pagingInfo = nextPaging;

            } while (!string.IsNullOrEmpty(pagingInfo));

            return new SharePointResult<List<DocumentReportItem>>
            {
                Data = documents,
                Status = SharePointResultStatus.Success,
                ErrorMessage = errors.Count > 0 ? string.Join("; ", errors) : null
            };
        }
        catch (HttpRequestException ex)
        {
            return new SharePointResult<List<DocumentReportItem>>
            {
                Data = documents, // Return what we have
                Status = ex.StatusCode switch
                {
                    HttpStatusCode.Unauthorized => SharePointResultStatus.AuthenticationRequired,
                    HttpStatusCode.Forbidden => SharePointResultStatus.AccessDenied,
                    HttpStatusCode.NotFound => SharePointResultStatus.NotFound,
                    _ => SharePointResultStatus.Error
                },
                ErrorMessage = $"{ex.Message}; Collected {documents.Count} docs before error"
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<List<DocumentReportItem>>
            {
                Data = documents, // Return what we have
                Status = SharePointResultStatus.Error,
                ErrorMessage = $"{ex.Message}; Collected {documents.Count} docs before error"
            };
        }
    }

    private static (List<DocumentReportItem> items, string? nextHref) ParseRenderListDataResponse(
        string json,
        string siteUrl,
        string libraryTitle)
    {
        var items = new List<DocumentReportItem>();
        string? nextHref = null;

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            // Get paging info
            if (root.TryGetProperty("NextHref", out var nextHrefElement) &&
                nextHrefElement.ValueKind == JsonValueKind.String)
            {
                nextHref = nextHrefElement.GetString();
            }

            // Get rows
            if (!root.TryGetProperty("Row", out var rowsElement))
            {
                return (items, nextHref);
            }

            var siteUri = new Uri(siteUrl);

            foreach (var row in rowsElement.EnumerateArray())
            {
                var fileName = GetStringProperty(row, "FileLeafRef");
                var fileRef = GetStringProperty(row, "FileRef");

                if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(fileRef))
                    continue;

                var extension = Path.GetExtension(fileName).TrimStart('.').ToLowerInvariant();
                var absoluteUrl = $"{siteUri.Scheme}://{siteUri.Host}{fileRef}";
                var folderPath = ExtractFolderPath(fileRef, libraryTitle, fileName);

                // Parse file size - try multiple field names as SharePoint can return different formats
                long fileSize = 0;
                var fileSizeStr = GetStringProperty(row, "File_x0020_Size");
                if (string.IsNullOrEmpty(fileSizeStr))
                {
                    fileSizeStr = GetStringProperty(row, "FileSizeDisplay");
                }
                if (string.IsNullOrEmpty(fileSizeStr))
                {
                    fileSizeStr = GetStringProperty(row, "SMTotalFileStreamSize");
                }
                if (!string.IsNullOrEmpty(fileSizeStr))
                {
                    // Remove any formatting (commas, spaces, "bytes" text)
                    var cleanSize = fileSizeStr.Replace(",", "").Replace(" ", "").Replace("bytes", "");
                    long.TryParse(cleanSize, out fileSize);
                }

                // Parse version (comes as string like "1.0" or "2.0")
                var versionStr = GetStringProperty(row, "_UIVersionString");
                int versionCount = 1;
                if (!string.IsNullOrEmpty(versionStr))
                {
                    var dotIndex = versionStr.IndexOf('.');
                    if (dotIndex > 0 && int.TryParse(versionStr.Substring(0, dotIndex), out var major))
                    {
                        versionCount = major;
                    }
                }

                // Author/Editor come as arrays with display text
                var createdBy = GetDisplayValue(row, "Author");
                var modifiedBy = GetDisplayValue(row, "Editor");

                items.Add(new DocumentReportItem
                {
                    FileName = fileName,
                    Extension = extension,
                    SizeBytes = fileSize,
                    CreatedDate = GetDateProperty(row, "Created"),
                    CreatedBy = createdBy,
                    ModifiedDate = GetDateProperty(row, "Modified"),
                    ModifiedBy = modifiedBy,
                    FileUrl = absoluteUrl,
                    ServerRelativeUrl = fileRef,
                    SiteCollectionUrl = siteUrl,
                    LibraryTitle = libraryTitle,
                    VersionCount = versionCount,
                    FolderPath = folderPath
                });
            }
        }
        catch
        {
            // Return what we have
        }

        return (items, nextHref);
    }

    private static string GetDisplayValue(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop))
        {
            // RenderListDataAsStream returns person fields as arrays
            if (prop.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in prop.EnumerateArray())
                {
                    if (item.TryGetProperty("title", out var title))
                    {
                        return title.GetString() ?? string.Empty;
                    }
                }
            }
            else if (prop.ValueKind == JsonValueKind.String)
            {
                return prop.GetString() ?? string.Empty;
            }
        }
        return string.Empty;
    }

    private static string ExtractFolderPath(string serverRelativeUrl, string libraryTitle, string fileName)
    {
        // ServerRelativeUrl format: /sites/sitename/libraryname/folder1/folder2/filename.ext
        // We want to extract: libraryname/folder1/folder2 (without the filename)

        try
        {
            // Remove the filename from the end
            var lastSlash = serverRelativeUrl.LastIndexOf('/');
            if (lastSlash <= 0)
                return string.Empty;

            var pathWithoutFile = serverRelativeUrl.Substring(0, lastSlash);

            // Find the library name in the path and extract everything from there
            var libraryIndex = pathWithoutFile.IndexOf("/" + libraryTitle, StringComparison.OrdinalIgnoreCase);
            if (libraryIndex >= 0)
            {
                // Return path starting from library name
                return pathWithoutFile.Substring(libraryIndex + 1);
            }

            // Fallback: return the last part of the path
            return pathWithoutFile;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static long GetLongProperty(JsonElement element, string propertyName)
    {
        if (element.TryGetProperty(propertyName, out var prop))
        {
            if (prop.ValueKind == JsonValueKind.Number)
                return prop.GetInt64();
            if (prop.ValueKind == JsonValueKind.String)
            {
                var str = prop.GetString();
                if (!string.IsNullOrEmpty(str) && long.TryParse(str, out var val))
                    return val;
            }
        }
        return 0;
    }

    public async Task<List<SubsiteInfo>> GetSubsitesAsync(string siteUrl)
    {
        var subsites = new List<SubsiteInfo>();

        try
        {
            var apiUrl = $"{siteUrl.TrimEnd('/')}/_api/web/webs?$select=Title,Url,ServerRelativeUrl,WebTemplate,Created,LastItemModifiedDate";
            var response = await _client.GetAsync(apiUrl);

            if (!response.IsSuccessStatusCode)
            {
                System.Diagnostics.Debug.WriteLine($"[SPOManager] GetSubsitesAsync - HTTP {(int)response.StatusCode} for {siteUrl}");
                return subsites;
            }

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            JsonElement valueElement;
            if (root.TryGetProperty("value", out valueElement))
            {
                // Standard OData
            }
            else if (root.TryGetProperty("d", out var dElement) &&
                     dElement.TryGetProperty("results", out valueElement))
            {
                // OData verbose
            }
            else
            {
                return subsites;
            }

            foreach (var item in valueElement.EnumerateArray())
            {
                subsites.Add(new SubsiteInfo
                {
                    Title = GetStringProperty(item, "Title"),
                    Url = GetStringProperty(item, "Url"),
                    ServerRelativeUrl = GetStringProperty(item, "ServerRelativeUrl"),
                    WebTemplate = GetStringProperty(item, "WebTemplate"),
                    Created = GetDateProperty(item, "Created"),
                    LastItemModifiedDate = GetDateProperty(item, "LastItemModifiedDate")
                });
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetSubsitesAsync - Exception: {ex.Message}");
        }

        return subsites;
    }

    #region Subsites Report Methods

    public async Task<SharePointResult<List<SubsiteReportItem>>> GetSubsitesForReportAsync(
        string siteUrl, string siteCollectionUrl)
    {
        try
        {
            var baseUrl = siteUrl.TrimEnd('/');
            var apiUrl = $"{baseUrl}/_api/web/webs?$select=Title,Url,ServerRelativeUrl,WebTemplate,Created,LastItemModifiedDate,Language";
            var response = await _client.GetAsync(apiUrl);

            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                response.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                return new SharePointResult<List<SubsiteReportItem>>
                {
                    Status = response.StatusCode == System.Net.HttpStatusCode.Unauthorized
                        ? SharePointResultStatus.AuthenticationRequired
                        : SharePointResultStatus.AccessDenied,
                    ErrorMessage = $"HTTP {(int)response.StatusCode}"
                };
            }

            if (!response.IsSuccessStatusCode)
            {
                return new SharePointResult<List<SubsiteReportItem>>
                {
                    Status = SharePointResultStatus.Error,
                    ErrorMessage = $"HTTP {(int)response.StatusCode}"
                };
            }

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            JsonElement valueElement;
            if (root.TryGetProperty("value", out valueElement))
            {
                // Standard OData
            }
            else if (root.TryGetProperty("d", out var dElement) &&
                     dElement.TryGetProperty("results", out valueElement))
            {
                // OData verbose
            }
            else
            {
                return new SharePointResult<List<SubsiteReportItem>>
                {
                    Data = [],
                    Status = SharePointResultStatus.Success
                };
            }

            // Get site title
            var siteTitle = "";
            try
            {
                var webUrl = $"{baseUrl}/_api/web?$select=Title";
                var webResponse = await _client.GetAsync(webUrl);
                if (webResponse.IsSuccessStatusCode)
                {
                    var webJson = await webResponse.Content.ReadAsStringAsync();
                    using var webDoc = JsonDocument.Parse(webJson);
                    var webRoot = webDoc.RootElement;
                    if (webRoot.TryGetProperty("d", out var dWeb))
                        siteTitle = GetStringProperty(dWeb, "Title");
                    else
                        siteTitle = GetStringProperty(webRoot, "Title");
                }
            }
            catch { /* ignore */ }

            var items = new List<SubsiteReportItem>();

            foreach (var item in valueElement.EnumerateArray())
            {
                var language = 0;
                if (item.TryGetProperty("Language", out var langProp))
                {
                    if (langProp.ValueKind == JsonValueKind.Number)
                        language = langProp.GetInt32();
                }

                items.Add(new SubsiteReportItem
                {
                    SiteCollectionUrl = siteCollectionUrl,
                    SiteUrl = siteUrl,
                    SiteTitle = siteTitle,
                    SubsiteUrl = GetStringProperty(item, "Url"),
                    SubsiteTitle = GetStringProperty(item, "Title"),
                    ServerRelativeUrl = GetStringProperty(item, "ServerRelativeUrl"),
                    WebTemplate = GetStringProperty(item, "WebTemplate"),
                    Created = GetDateProperty(item, "Created"),
                    LastModified = GetDateProperty(item, "LastItemModifiedDate"),
                    Language = language
                });
            }

            return new SharePointResult<List<SubsiteReportItem>>
            {
                Data = items,
                Status = SharePointResultStatus.Success
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<List<SubsiteReportItem>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    #endregion

    #region Publishing Feature Methods

    public async Task<SharePointResult<SitePublishingResult>> GetPublishingFeatureStatusAsync(string siteUrl)
    {
        try
        {
            var baseUrl = siteUrl.TrimEnd('/');
            var result = new SitePublishingResult { SiteUrl = siteUrl };

            // Get site title
            try
            {
                var siteInfo = await GetSiteInfoAsync(siteUrl);
                result.SiteTitle = siteInfo.Title;
            }
            catch { /* ignore */ }

            // Check site collection features for Publishing Infrastructure
            var siteFeatureIds = new HashSet<Guid>();
            var siteFeaturesUrl = $"{baseUrl}/_api/site/features?$select=DefinitionId";
            var siteFeaturesResponse = await _client.GetAsync(siteFeaturesUrl);

            if (siteFeaturesResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                siteFeaturesResponse.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                return new SharePointResult<SitePublishingResult>
                {
                    Status = GetSharePointStatus(siteFeaturesResponse.StatusCode),
                    ErrorMessage = GetErrorMessage(siteFeaturesResponse.StatusCode)
                };
            }

            if (siteFeaturesResponse.IsSuccessStatusCode)
            {
                var json = await siteFeaturesResponse.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                JsonElement valueElement;
                if (root.TryGetProperty("value", out valueElement) ||
                    (root.TryGetProperty("d", out var d) && d.TryGetProperty("results", out valueElement)))
                {
                    foreach (var feature in valueElement.EnumerateArray())
                    {
                        var defId = GetGuidProperty(feature, "DefinitionId");
                        if (defId != Guid.Empty)
                            siteFeatureIds.Add(defId);
                    }
                }
            }

            result.HasPublishingInfrastructure = siteFeatureIds.Contains(PublishingFeatureIds.PublishingInfrastructure);

            // Check web features for Publishing
            var webFeatureIds = new HashSet<Guid>();
            var webFeaturesUrl = $"{baseUrl}/_api/web/features?$select=DefinitionId";
            var webFeaturesResponse = await _client.GetAsync(webFeaturesUrl);

            if (webFeaturesResponse.IsSuccessStatusCode)
            {
                var json = await webFeaturesResponse.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                JsonElement valueElement;
                if (root.TryGetProperty("value", out valueElement) ||
                    (root.TryGetProperty("d", out var d) && d.TryGetProperty("results", out valueElement)))
                {
                    foreach (var feature in valueElement.EnumerateArray())
                    {
                        var defId = GetGuidProperty(feature, "DefinitionId");
                        if (defId != Guid.Empty)
                            webFeatureIds.Add(defId);
                    }
                }
            }

            result.HasPublishingWeb = webFeatureIds.Contains(PublishingFeatureIds.PublishingWeb);
            result.Success = true;

            return new SharePointResult<SitePublishingResult>
            {
                Data = result,
                Status = SharePointResultStatus.Success
            };
        }
        catch (HttpRequestException ex) when (ex.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                                                ex.StatusCode == System.Net.HttpStatusCode.Forbidden)
        {
            return new SharePointResult<SitePublishingResult>
            {
                Status = ex.StatusCode == System.Net.HttpStatusCode.Unauthorized
                    ? SharePointResultStatus.AuthenticationRequired
                    : SharePointResultStatus.AccessDenied,
                ErrorMessage = ex.Message
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<SitePublishingResult>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    #endregion

    #region List Customization Methods

    public async Task<SharePointResult<List<CustomizedListItem>>> GetListFormCustomizationsAsync(string siteUrl)
    {
        try
        {
            var items = new List<CustomizedListItem>();
            var baseUrl = siteUrl.TrimEnd('/');

            // Get site title
            var siteTitle = string.Empty;
            try
            {
                var siteInfo = await GetSiteInfoAsync(siteUrl);
                siteTitle = siteInfo.Title;
            }
            catch { /* ignore */ }

            // Get all visible lists with form URLs
            var listsUrl = $"{baseUrl}/_api/web/lists?$select=Id,Title,BaseType,Hidden,ItemCount,DefaultNewFormUrl,DefaultEditFormUrl,DefaultDisplayFormUrl,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false";
            var response = await _client.GetAsync(listsUrl);

            if (!response.IsSuccessStatusCode)
            {
                return new SharePointResult<List<CustomizedListItem>>
                {
                    Status = GetSharePointStatus(response.StatusCode),
                    ErrorMessage = GetErrorMessage(response.StatusCode)
                };
            }

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            JsonElement valueElement;
            if (root.TryGetProperty("value", out valueElement))
            {
                // Standard OData nometadata
            }
            else if (root.TryGetProperty("d", out var dElement) &&
                     dElement.TryGetProperty("results", out valueElement))
            {
                // OData verbose
            }
            else
            {
                return new SharePointResult<List<CustomizedListItem>>
                {
                    Data = items,
                    Status = SharePointResultStatus.Success
                };
            }

            foreach (var listElement in valueElement.EnumerateArray())
            {
                var baseType = GetIntProperty(listElement, "BaseType");
                // Only GenericList (0) and DocumentLibrary (1)
                if (baseType != 0 && baseType != 1) continue;

                var listId = GetGuidProperty(listElement, "Id");
                var listTitle = GetStringProperty(listElement, "Title");
                var itemCount = GetIntProperty(listElement, "ItemCount");
                var newFormUrl = GetStringProperty(listElement, "DefaultNewFormUrl");
                var editFormUrl = GetStringProperty(listElement, "DefaultEditFormUrl");
                var displayFormUrl = GetStringProperty(listElement, "DefaultDisplayFormUrl");

                var serverRelativeUrl = string.Empty;
                if (listElement.TryGetProperty("RootFolder", out var rf))
                {
                    serverRelativeUrl = GetStringProperty(rf, "ServerRelativeUrl");
                }

                var formType = ListFormType.Default;
                var spfxNewId = string.Empty;
                var spfxEditId = string.Empty;
                var spfxDisplayId = string.Empty;

                // Check 1: URL-based Power Apps detection (quick)
                if (newFormUrl.Contains("PowerApps", StringComparison.OrdinalIgnoreCase) ||
                    editFormUrl.Contains("PowerApps", StringComparison.OrdinalIgnoreCase))
                {
                    formType = ListFormType.PowerApps;
                }

                // Check 2: Property bag Power Apps detection
                if (formType == ListFormType.Default && listId != Guid.Empty)
                {
                    try
                    {
                        var propsUrl = $"{baseUrl}/_api/web/lists(guid'{listId}')/RootFolder/Properties";
                        var propsResponse = await _client.GetAsync(propsUrl);
                        if (propsResponse.IsSuccessStatusCode)
                        {
                            var propsJson = await propsResponse.Content.ReadAsStringAsync();
                            if (propsJson.Contains("PowerAppsFormProperties", StringComparison.OrdinalIgnoreCase) ||
                                propsJson.Contains("PowerAppFormProperties", StringComparison.OrdinalIgnoreCase) ||
                                propsJson.Contains("_PowerAppsId_", StringComparison.OrdinalIgnoreCase) ||
                                propsJson.Contains("PowerAppsFormId", StringComparison.OrdinalIgnoreCase))
                            {
                                formType = ListFormType.PowerApps;
                            }
                        }
                    }
                    catch { /* skip property bag check */ }
                }

                // Check 3: SPFx Content Type detection
                if (formType == ListFormType.Default && listId != Guid.Empty)
                {
                    try
                    {
                        var ctUrl = $"{baseUrl}/_api/web/lists(guid'{listId}')/ContentTypes?$select=Name,StringId,NewFormClientSideComponentId,EditFormClientSideComponentId,DisplayFormClientSideComponentId";
                        var ctResponse = await _client.GetAsync(ctUrl);
                        if (ctResponse.IsSuccessStatusCode)
                        {
                            var ctJson = await ctResponse.Content.ReadAsStringAsync();
                            using var ctDoc = JsonDocument.Parse(ctJson);
                            var ctRoot = ctDoc.RootElement;

                            JsonElement ctValue;
                            if (ctRoot.TryGetProperty("value", out ctValue) ||
                                (ctRoot.TryGetProperty("d", out var ctD) && ctD.TryGetProperty("results", out ctValue)))
                            {
                                foreach (var ct in ctValue.EnumerateArray())
                                {
                                    var newComp = GetStringProperty(ct, "NewFormClientSideComponentId");
                                    var editComp = GetStringProperty(ct, "EditFormClientSideComponentId");
                                    var displayComp = GetStringProperty(ct, "DisplayFormClientSideComponentId");

                                    // Non-empty and not the empty GUID means SPFx is configured
                                    var hasSpfx = IsNonEmptyGuid(newComp) || IsNonEmptyGuid(editComp) || IsNonEmptyGuid(displayComp);
                                    if (hasSpfx)
                                    {
                                        formType = ListFormType.SPFxCustomForm;
                                        spfxNewId = IsNonEmptyGuid(newComp) ? newComp : string.Empty;
                                        spfxEditId = IsNonEmptyGuid(editComp) ? editComp : string.Empty;
                                        spfxDisplayId = IsNonEmptyGuid(displayComp) ? displayComp : string.Empty;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    catch { /* skip SPFx check */ }
                }

                items.Add(new CustomizedListItem
                {
                    SiteUrl = siteUrl,
                    SiteTitle = siteTitle,
                    ListId = listId,
                    ListTitle = listTitle,
                    ListType = baseType == 0 ? "List" : "Document Library",
                    FormType = formType,
                    ItemCount = itemCount,
                    ListUrl = serverRelativeUrl,
                    DefaultNewFormUrl = newFormUrl,
                    DefaultEditFormUrl = editFormUrl,
                    DefaultDisplayFormUrl = displayFormUrl,
                    SpfxNewFormComponentId = spfxNewId,
                    SpfxEditFormComponentId = spfxEditId,
                    SpfxDisplayFormComponentId = spfxDisplayId
                });
            }

            return new SharePointResult<List<CustomizedListItem>>
            {
                Data = items,
                Status = SharePointResultStatus.Success
            };
        }
        catch (HttpRequestException ex) when (ex.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                                                ex.StatusCode == System.Net.HttpStatusCode.Forbidden)
        {
            return new SharePointResult<List<CustomizedListItem>>
            {
                Status = ex.StatusCode == System.Net.HttpStatusCode.Unauthorized
                    ? SharePointResultStatus.AuthenticationRequired
                    : SharePointResultStatus.AccessDenied,
                ErrorMessage = ex.Message
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<List<CustomizedListItem>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    private static bool IsNonEmptyGuid(string value)
    {
        return !string.IsNullOrEmpty(value) &&
               Guid.TryParse(value, out var guid) &&
               guid != Guid.Empty;
    }

    #endregion

    #region Custom Fields Methods

    public async Task<SharePointResult<List<CustomFieldItem>>> GetListCustomFieldsAsync(
        string siteUrl, string siteCollectionUrl)
    {
        try
        {
            var items = new List<CustomFieldItem>();
            var baseUrl = siteUrl.TrimEnd('/');

            // Get site title
            var siteTitle = string.Empty;
            try
            {
                var siteInfo = await GetSiteInfoAsync(siteUrl);
                siteTitle = siteInfo.Title;
            }
            catch { /* ignore */ }

            // Get all visible lists with metadata
            var listsUrl = $"{baseUrl}/_api/web/lists?$select=Id,Title,BaseType,Hidden,ItemCount,Created,LastItemModifiedDate,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false";
            var response = await _client.GetAsync(listsUrl);

            if (!response.IsSuccessStatusCode)
            {
                return new SharePointResult<List<CustomFieldItem>>
                {
                    Status = GetSharePointStatus(response.StatusCode),
                    ErrorMessage = GetErrorMessage(response.StatusCode)
                };
            }

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            JsonElement valueElement;
            if (root.TryGetProperty("value", out valueElement))
            {
                // Standard OData nometadata
            }
            else if (root.TryGetProperty("d", out var dElement) &&
                     dElement.TryGetProperty("results", out valueElement))
            {
                // OData verbose
            }
            else
            {
                return new SharePointResult<List<CustomFieldItem>>
                {
                    Data = items,
                    Status = SharePointResultStatus.Success
                };
            }

            foreach (var listElement in valueElement.EnumerateArray())
            {
                var baseType = GetIntProperty(listElement, "BaseType");
                // Only GenericList (0) and DocumentLibrary (1)
                if (baseType != 0 && baseType != 1) continue;

                var listId = GetGuidProperty(listElement, "Id");
                var listTitle = GetStringProperty(listElement, "Title");
                var itemCount = GetIntProperty(listElement, "ItemCount");

                var listCreated = DateTime.MinValue;
                var listModified = DateTime.MinValue;
                if (listElement.TryGetProperty("Created", out var createdProp))
                {
                    DateTime.TryParse(createdProp.GetString(), out listCreated);
                }
                if (listElement.TryGetProperty("LastItemModifiedDate", out var modifiedProp))
                {
                    DateTime.TryParse(modifiedProp.GetString(), out listModified);
                }

                var serverRelativeUrl = string.Empty;
                if (listElement.TryGetProperty("RootFolder", out var rf))
                {
                    serverRelativeUrl = GetStringProperty(rf, "ServerRelativeUrl");
                }

                if (listId == Guid.Empty) continue;

                // Get fields for this list
                try
                {
                    var fieldsUrl = $"{baseUrl}/_api/web/lists(guid'{listId}')/fields?$filter=Hidden eq false and FromBaseType eq false&$select=Title,InternalName,TypeAsString,Group,CanBeDeleted,FromBaseType";
                    var fieldsResponse = await _client.GetAsync(fieldsUrl);

                    if (!fieldsResponse.IsSuccessStatusCode)
                    {
                        // Fallback: some environments may not support FromBaseType in $filter
                        fieldsUrl = $"{baseUrl}/_api/web/lists(guid'{listId}')/fields?$filter=Hidden eq false&$select=Title,InternalName,TypeAsString,Group,CanBeDeleted,FromBaseType";
                        fieldsResponse = await _client.GetAsync(fieldsUrl);
                        if (!fieldsResponse.IsSuccessStatusCode) continue;
                    }

                    var fieldsJson = await fieldsResponse.Content.ReadAsStringAsync();
                    using var fieldsDoc = JsonDocument.Parse(fieldsJson);
                    var fieldsRoot = fieldsDoc.RootElement;

                    JsonElement fieldsValue;
                    if (fieldsRoot.TryGetProperty("value", out fieldsValue))
                    {
                        // Standard
                    }
                    else if (fieldsRoot.TryGetProperty("d", out var fd) &&
                             fd.TryGetProperty("results", out fieldsValue))
                    {
                        // Verbose
                    }
                    else
                    {
                        continue;
                    }

                    foreach (var fieldElement in fieldsValue.EnumerateArray())
                    {
                        var group = GetStringProperty(fieldElement, "Group");

                        if (!OotbFieldGroups.IsCustomGroup(group)) continue;

                        // Skip system fields that SharePoint puts in "Custom Columns"
                        var fromBaseType = GetBoolProperty(fieldElement, "FromBaseType");
                        if (fromBaseType) continue;

                        var internalName = GetStringProperty(fieldElement, "InternalName");
                        if (OotbFieldGroups.IsSystemFieldInternalName(internalName)) continue;

                        // OOTB list-template fields typically cannot be deleted
                        var canBeDeleted = GetBoolProperty(fieldElement, "CanBeDeleted");
                        if (!canBeDeleted) continue;

                        items.Add(new CustomFieldItem
                        {
                            SiteCollectionUrl = siteCollectionUrl,
                            SiteUrl = siteUrl,
                            SiteTitle = siteTitle,
                            ListTitle = listTitle,
                            ListUrl = serverRelativeUrl,
                            ItemCount = itemCount,
                            ColumnName = GetStringProperty(fieldElement, "Title"),
                            InternalName = internalName,
                            FieldType = GetStringProperty(fieldElement, "TypeAsString"),
                            Group = group,
                            ListCreated = listCreated,
                            ListModified = listModified
                        });
                    }
                }
                catch
                {
                    // Skip lists whose fields can't be read
                }
            }

            return new SharePointResult<List<CustomFieldItem>>
            {
                Data = items,
                Status = SharePointResultStatus.Success
            };
        }
        catch (HttpRequestException ex) when (ex.StatusCode == HttpStatusCode.Unauthorized ||
                                                ex.StatusCode == HttpStatusCode.Forbidden)
        {
            return new SharePointResult<List<CustomFieldItem>>
            {
                Status = SharePointResultStatus.AuthenticationRequired,
                ErrorMessage = "Authentication required"
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<List<CustomFieldItem>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    #endregion

    #region Site Users Methods

    public async Task<SharePointResult<List<AdHocUserItem>>> GetSiteUsersAsync(
        string siteUrl, string? loginNameFilter = null)
    {
        try
        {
            var users = new List<AdHocUserItem>();
            var baseUrl = siteUrl.TrimEnd('/');

            // Try server-side filter first (double-encode %3a → %253a for OData filter on URL-encoded LoginName)
            bool useClientFilter = false;
            string apiUrl;

            if (!string.IsNullOrEmpty(loginNameFilter))
            {
                var encodedFilter = loginNameFilter.Replace("%3a", "%253a");
                apiUrl = $"{baseUrl}/_api/web/siteusers?$select=Id,LoginName,Title,Email,IsSiteAdmin,PrincipalType&$filter=substringof('{encodedFilter}',LoginName)";
            }
            else
            {
                apiUrl = $"{baseUrl}/_api/web/siteusers?$select=Id,LoginName,Title,Email,IsSiteAdmin,PrincipalType";
            }

            var response = await _client.GetAsync(apiUrl);

            // If server-side filter fails, fall back to fetching all users and filtering client-side
            if (!response.IsSuccessStatusCode && !string.IsNullOrEmpty(loginNameFilter))
            {
                useClientFilter = true;
                apiUrl = $"{baseUrl}/_api/web/siteusers?$select=Id,LoginName,Title,Email,IsSiteAdmin,PrincipalType";
                response = await _client.GetAsync(apiUrl);
            }

            if (!response.IsSuccessStatusCode)
            {
                return new SharePointResult<List<AdHocUserItem>>
                {
                    Status = GetSharePointStatus(response.StatusCode),
                    ErrorMessage = GetErrorMessage(response.StatusCode)
                };
            }

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            JsonElement valueElement;
            if (root.TryGetProperty("value", out valueElement))
            {
                // Standard OData nometadata
            }
            else if (root.TryGetProperty("d", out var dElement) &&
                     dElement.TryGetProperty("results", out valueElement))
            {
                // OData verbose
            }
            else
            {
                return new SharePointResult<List<AdHocUserItem>>
                {
                    Data = users,
                    Status = SharePointResultStatus.Success
                };
            }

            // Get site title for context
            var siteTitle = string.Empty;
            try
            {
                var siteInfo = await GetSiteInfoAsync(siteUrl);
                siteTitle = siteInfo.Title;
            }
            catch { /* ignore */ }

            foreach (var item in valueElement.EnumerateArray())
            {
                var loginName = GetStringProperty(item, "LoginName");

                // Client-side filter if server-side filter wasn't used
                if (useClientFilter && !string.IsNullOrEmpty(loginNameFilter))
                {
                    if (!loginName.Contains(loginNameFilter, StringComparison.OrdinalIgnoreCase))
                        continue;
                }

                var principalTypeInt = GetIntProperty(item, "PrincipalType");
                var principalTypeStr = principalTypeInt switch
                {
                    1 => "User",
                    2 => "DistributionList",
                    4 => "SecurityGroup",
                    8 => "SharePointGroup",
                    _ => principalTypeInt.ToString()
                };

                users.Add(new AdHocUserItem
                {
                    SiteUrl = siteUrl,
                    SiteTitle = siteTitle,
                    LoginName = loginName,
                    Title = GetStringProperty(item, "Title"),
                    Email = GetStringProperty(item, "Email"),
                    Id = GetIntProperty(item, "Id"),
                    IsSiteAdmin = GetBoolProperty(item, "IsSiteAdmin"),
                    PrincipalType = principalTypeStr
                });
            }

            return new SharePointResult<List<AdHocUserItem>>
            {
                Data = users,
                Status = SharePointResultStatus.Success
            };
        }
        catch (HttpRequestException ex) when (ex.StatusCode == System.Net.HttpStatusCode.Unauthorized ||
                                                ex.StatusCode == System.Net.HttpStatusCode.Forbidden)
        {
            return new SharePointResult<List<AdHocUserItem>>
            {
                Status = ex.StatusCode == System.Net.HttpStatusCode.Unauthorized
                    ? SharePointResultStatus.AuthenticationRequired
                    : SharePointResultStatus.AccessDenied,
                ErrorMessage = ex.Message
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<List<AdHocUserItem>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    #endregion

    #region Permission Methods

    public async Task<SharePointResult<List<PermissionReportItem>>> GetWebPermissionsAsync(
        string siteUrl,
        string siteCollectionUrl,
        bool includeInherited = false)
    {
        System.Diagnostics.Debug.WriteLine($"[SPOManager] GetWebPermissionsAsync - START");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   siteUrl: {siteUrl}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   siteCollectionUrl: {siteCollectionUrl}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   includeInherited: {includeInherited}");

        try
        {
            var permissions = new List<PermissionReportItem>();
            var baseUrl = siteUrl.TrimEnd('/');

            // Check if web has unique permissions
            var hasUniqueUrl = $"{baseUrl}/_api/web?$select=Title,Url,HasUniqueRoleAssignments,ServerRelativeUrl";
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Calling: {hasUniqueUrl}");

            var webResponse = await _client.GetAsync(hasUniqueUrl);
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Response: {(int)webResponse.StatusCode} {webResponse.StatusCode}");

            if (!webResponse.IsSuccessStatusCode)
            {
                var errorContent = await webResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   ERROR Response Body: {errorContent.Substring(0, Math.Min(500, errorContent.Length))}");
                return new SharePointResult<List<PermissionReportItem>>
                {
                    Status = GetSharePointStatus(webResponse.StatusCode),
                    ErrorMessage = GetErrorMessage(webResponse.StatusCode)
                };
            }

            var webJson = await webResponse.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Web JSON (first 500 chars): {webJson.Substring(0, Math.Min(500, webJson.Length))}");

            using var webDoc = JsonDocument.Parse(webJson);
            var webRoot = webDoc.RootElement;

            var webTitle = GetStringProperty(webRoot, "Title");
            var webUrl = GetStringProperty(webRoot, "Url");
            var serverRelativeUrl = GetStringProperty(webRoot, "ServerRelativeUrl");
            var hasUniquePerms = GetBoolProperty(webRoot, "HasUniqueRoleAssignments");

            System.Diagnostics.Debug.WriteLine($"[SPOManager]   webTitle: {webTitle}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   webUrl: {webUrl}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   serverRelativeUrl: {serverRelativeUrl}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   hasUniquePerms: {hasUniquePerms}");

            // Determine object type
            var objectType = siteUrl.Equals(siteCollectionUrl, StringComparison.OrdinalIgnoreCase)
                ? PermissionObjectType.SiteCollection
                : (serverRelativeUrl.Count(c => c == '/') > 2 ? PermissionObjectType.Subsite : PermissionObjectType.Site);

            System.Diagnostics.Debug.WriteLine($"[SPOManager]   objectType: {objectType}");

            // If inherited and we don't want inherited, skip
            if (!hasUniquePerms && !includeInherited)
            {
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Skipping - inherited permissions and includeInherited=false");
                return new SharePointResult<List<PermissionReportItem>>
                {
                    Data = permissions,
                    Status = SharePointResultStatus.Success
                };
            }

            // Get role assignments
            var roleUrl = $"{baseUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings";
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Calling role assignments: {roleUrl}");

            var roleResponse = await _client.GetAsync(roleUrl);
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Role response: {(int)roleResponse.StatusCode} {roleResponse.StatusCode}");

            if (roleResponse.IsSuccessStatusCode)
            {
                var roleJson = await roleResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Role JSON length: {roleJson.Length}");
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Role JSON (first 1000 chars): {roleJson.Substring(0, Math.Min(1000, roleJson.Length))}");

                var roleItems = ParseRoleAssignments(roleJson, siteCollectionUrl, siteUrl, webTitle,
                    objectType, webTitle, webUrl, !hasUniquePerms, siteUrl);
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Parsed {roleItems.Count} permission entries");
                permissions.AddRange(roleItems);
            }
            else
            {
                var errorContent = await roleResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Role ERROR Response: {errorContent.Substring(0, Math.Min(500, errorContent.Length))}");
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetWebPermissionsAsync - END - returning {permissions.Count} permissions");
            return new SharePointResult<List<PermissionReportItem>>
            {
                Data = permissions,
                Status = SharePointResultStatus.Success
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetWebPermissionsAsync - EXCEPTION: {ex.GetType().Name}: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   StackTrace: {ex.StackTrace}");
            return new SharePointResult<List<PermissionReportItem>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    public async Task<SharePointResult<List<PermissionReportItem>>> GetListPermissionsAsync(
        string siteUrl,
        string siteCollectionUrl,
        string listTitle,
        bool isLibrary,
        bool includeInherited = false)
    {
        System.Diagnostics.Debug.WriteLine($"[SPOManager] GetListPermissionsAsync - START");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   siteUrl: {siteUrl}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   listTitle: {listTitle}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   isLibrary: {isLibrary}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   includeInherited: {includeInherited}");

        try
        {
            var permissions = new List<PermissionReportItem>();
            var baseUrl = siteUrl.TrimEnd('/');
            var encodedListTitle = Uri.EscapeDataString(listTitle);

            // Get list info including HasUniqueRoleAssignments
            var listUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedListTitle}')?$select=Title,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl&$expand=RootFolder";
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Calling: {listUrl}");

            var listResponse = await _client.GetAsync(listUrl);
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Response: {(int)listResponse.StatusCode} {listResponse.StatusCode}");

            if (!listResponse.IsSuccessStatusCode)
            {
                var errorContent = await listResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   ERROR Response: {errorContent.Substring(0, Math.Min(500, errorContent.Length))}");
                return new SharePointResult<List<PermissionReportItem>>
                {
                    Status = GetSharePointStatus(listResponse.StatusCode),
                    ErrorMessage = GetErrorMessage(listResponse.StatusCode)
                };
            }

            var listJson = await listResponse.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   List JSON (first 500 chars): {listJson.Substring(0, Math.Min(500, listJson.Length))}");

            using var listDoc = JsonDocument.Parse(listJson);
            var listRoot = listDoc.RootElement;

            var hasUniquePerms = GetBoolProperty(listRoot, "HasUniqueRoleAssignments");
            var listServerRelativeUrl = string.Empty;
            if (listRoot.TryGetProperty("RootFolder", out var rootFolder))
            {
                listServerRelativeUrl = GetStringProperty(rootFolder, "ServerRelativeUrl");
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager]   hasUniquePerms: {hasUniquePerms}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   listServerRelativeUrl: {listServerRelativeUrl}");

            var siteUri = new Uri(siteUrl);
            var listAbsoluteUrl = !string.IsNullOrEmpty(listServerRelativeUrl)
                ? $"{siteUri.Scheme}://{siteUri.Host}{listServerRelativeUrl}"
                : siteUrl;

            // If inherited and we don't want inherited, skip
            if (!hasUniquePerms && !includeInherited)
            {
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Skipping list - inherited permissions and includeInherited=false");
                return new SharePointResult<List<PermissionReportItem>>
                {
                    Data = permissions,
                    Status = SharePointResultStatus.Success
                };
            }

            // Get site title for context
            var siteInfo = await GetSiteInfoAsync(siteUrl);
            var siteTitle = siteInfo.Title;

            var objectType = isLibrary ? PermissionObjectType.Library : PermissionObjectType.List;

            // Get role assignments
            var roleUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedListTitle}')/roleassignments?$expand=Member,RoleDefinitionBindings";
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Calling role assignments: {roleUrl}");

            var roleResponse = await _client.GetAsync(roleUrl);
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Role response: {(int)roleResponse.StatusCode} {roleResponse.StatusCode}");

            if (roleResponse.IsSuccessStatusCode)
            {
                var roleJson = await roleResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Role JSON length: {roleJson.Length}");

                var roleItems = ParseRoleAssignments(roleJson, siteCollectionUrl, siteUrl, siteTitle,
                    objectType, listTitle, listAbsoluteUrl, !hasUniquePerms, siteUrl);
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Parsed {roleItems.Count} permission entries for list");
                permissions.AddRange(roleItems);
            }
            else
            {
                var errorContent = await roleResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Role ERROR: {errorContent.Substring(0, Math.Min(500, errorContent.Length))}");
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetListPermissionsAsync - END - returning {permissions.Count} permissions");
            return new SharePointResult<List<PermissionReportItem>>
            {
                Data = permissions,
                Status = SharePointResultStatus.Success
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetListPermissionsAsync - EXCEPTION: {ex.GetType().Name}: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   StackTrace: {ex.StackTrace}");
            return new SharePointResult<List<PermissionReportItem>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    public async Task<SharePointResult<List<PermissionReportItem>>> GetItemPermissionsAsync(
        string siteUrl,
        string siteCollectionUrl,
        string listTitle,
        bool isLibrary,
        bool includeFolders = true,
        bool includeItems = true,
        bool includeInherited = false,
        Action<int, int>? onPageScanned = null)
    {
        System.Diagnostics.Debug.WriteLine($"[SPOManager] GetItemPermissionsAsync - START");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   siteUrl: {siteUrl}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   listTitle: {listTitle}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   isLibrary: {isLibrary}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   includeFolders: {includeFolders}, includeItems: {includeItems}");

        try
        {
            var permissions = new List<PermissionReportItem>();
            var baseUrl = siteUrl.TrimEnd('/');
            var encodedListTitle = Uri.EscapeDataString(listTitle);

            // Get site title
            var siteInfo = await GetSiteInfoAsync(siteUrl);
            var siteTitle = siteInfo.Title;

            var siteUri = new Uri(siteUrl);

            // HasUniqueRoleAssignments is NOT reliably returned by the REST API items endpoint.
            // We use CSOM ProcessQuery (like Sharegate) to batch-check this property reliably.
            var formDigest = await GetFormDigestValueAsync(baseUrl);
            var selectFields = "Id,FileRef,FileLeafRef,Title,FileSystemObjectType";
            int lastId = 0;
            bool hasMore = true;
            int pageNumber = 0;
            int totalItemsScanned = 0;
            int itemsWithUniquePerms = 0;

            while (hasMore)
            {
                pageNumber++;
                var apiUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedListTitle}')/items" +
                            $"?$select={selectFields}&$top=5000&$orderby=Id asc&$filter=Id gt {lastId}";

                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Page {pageNumber}: {apiUrl}");

                var response = await _client.GetAsync(apiUrl);
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Page {pageNumber} response: {(int)response.StatusCode} {response.StatusCode}");

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"[SPOManager]   ERROR: {errorContent.Substring(0, Math.Min(500, errorContent.Length))}");
                    return new SharePointResult<List<PermissionReportItem>>
                    {
                        Status = GetSharePointStatus(response.StatusCode),
                        ErrorMessage = GetErrorMessage(response.StatusCode)
                    };
                }

                var json = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                JsonElement valueElement;
                if (root.TryGetProperty("value", out valueElement))
                {
                    // Standard OData
                }
                else if (root.TryGetProperty("d", out var dElement) &&
                         dElement.TryGetProperty("results", out valueElement))
                {
                    // OData verbose
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[SPOManager]   Could not find 'value' or 'd.results' in response");
                    break;
                }

                // Collect all items from this page
                var pageItems = new List<(int Id, string FileRef, string FileName, int FsObjType)>();
                foreach (var item in valueElement.EnumerateArray())
                {
                    totalItemsScanned++;
                    var itemId = GetIntProperty(item, "Id");
                    lastId = itemId;

                    var fileRef = GetStringProperty(item, "FileRef");
                    var fileName = GetStringProperty(item, "FileLeafRef");
                    var fsObjType = GetIntProperty(item, "FileSystemObjectType");
                    if (string.IsNullOrEmpty(fileName))
                        fileName = GetStringProperty(item, "Title");

                    pageItems.Add((itemId, fileRef, fileName, fsObjType));
                }

                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Page {pageNumber}: {pageItems.Count} items, lastId={lastId}");

                // Use CSOM ProcessQuery to batch-check HasUniqueRoleAssignments.
                // CSOM is more reliable than REST for detecting file-level unique permissions.
                var allItemIds = pageItems.Select(pi => pi.Id).ToList();
                var uniqueSet = await CheckUniquePermissionsCsomAsync(baseUrl, formDigest, listTitle, allItemIds);

                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Page {pageNumber}: {uniqueSet.Count} items with unique permissions out of {pageItems.Count}");
                foreach (var (itemId, fileRef, fileName, fsObjType) in pageItems)
                {
                    if (!uniqueSet.Contains(itemId))
                        continue;

                    itemsWithUniquePerms++;
                    var isFolder = fsObjType == 1;

                    System.Diagnostics.Debug.WriteLine($"[SPOManager]     Item {itemId}: hasUnique=true, isFolder={isFolder}, name={fileName}");

                    // Skip based on options
                    if (isFolder && !includeFolders)
                    {
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]       Skipping folder (includeFolders=false)");
                        continue;
                    }
                    if (!isFolder && !includeItems)
                    {
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]       Skipping item (includeItems=false)");
                        continue;
                    }

                    var absoluteUrl = !string.IsNullOrEmpty(fileRef)
                        ? $"{siteUri.Scheme}://{siteUri.Host}{fileRef}"
                        : siteUrl;
                    var objectType = isFolder ? PermissionObjectType.Folder
                        : (isLibrary ? PermissionObjectType.Document : PermissionObjectType.ListItem);

                    // Get role assignments for this item
                    var roleUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedListTitle}')/items({itemId})/roleassignments?$expand=Member,RoleDefinitionBindings";
                    System.Diagnostics.Debug.WriteLine($"[SPOManager]       Fetching role assignments for item {itemId}");

                    var roleResponse = await _client.GetAsync(roleUrl);

                    if (roleResponse.IsSuccessStatusCode)
                    {
                        var roleJson = await roleResponse.Content.ReadAsStringAsync();
                        var roleItems = ParseRoleAssignments(roleJson, siteCollectionUrl, siteUrl, siteTitle,
                            objectType, fileName, absoluteUrl, false, siteUrl);

                        System.Diagnostics.Debug.WriteLine($"[SPOManager]       Found {roleItems.Count} role assignments for item {itemId}");

                        foreach (var roleItem in roleItems)
                        {
                            roleItem.ObjectPath = fileRef;
                        }

                        permissions.AddRange(roleItems);
                    }
                    else
                    {
                        var errorContent = await roleResponse.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]       Role assignment ERROR for item {itemId}: {(int)roleResponse.StatusCode}");
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]       Error: {errorContent.Substring(0, Math.Min(300, errorContent.Length))}");
                    }
                }

                onPageScanned?.Invoke(totalItemsScanned, itemsWithUniquePerms);

                // If we got fewer than 5000 items, we've reached the end
                hasMore = pageItems.Count == 5000;
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetItemPermissionsAsync - END");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Total items scanned: {totalItemsScanned}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Items with unique permissions: {itemsWithUniquePerms}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Total permission entries: {permissions.Count}");

            return new SharePointResult<List<PermissionReportItem>>
            {
                Data = permissions,
                Status = SharePointResultStatus.Success
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetItemPermissionsAsync - EXCEPTION: {ex.GetType().Name}: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   StackTrace: {ex.StackTrace}");
            return new SharePointResult<List<PermissionReportItem>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    /// <summary>
    /// Gets a form digest value needed for CSOM POST requests.
    /// </summary>
    private async Task<string?> GetFormDigestValueAsync(string baseUrl)
    {
        Trace.WriteLine($"[SPOManager] CSOM: Getting form digest from {baseUrl}/_api/contextinfo");
        var digestUrl = $"{baseUrl}/_api/contextinfo";
        var digestResponse = await _client.PostAsync(digestUrl, new StringContent(""));

        if (!digestResponse.IsSuccessStatusCode)
        {
            Trace.WriteLine($"[SPOManager] CSOM: GetFormDigest FAILED HTTP {(int)digestResponse.StatusCode}");
            return null;
        }

        var digestJson = await digestResponse.Content.ReadAsStringAsync();
        using var digestDoc = JsonDocument.Parse(digestJson);

        var formDigestValue = GetStringProperty(digestDoc.RootElement, "FormDigestValue");
        if (string.IsNullOrEmpty(formDigestValue) &&
            digestDoc.RootElement.TryGetProperty("d", out var dElement) &&
            dElement.TryGetProperty("GetContextWebInformation", out var webInfo))
        {
            formDigestValue = GetStringProperty(webInfo, "FormDigestValue");
        }

        Trace.WriteLine($"[SPOManager] CSOM: Got form digest: {(string.IsNullOrEmpty(formDigestValue) ? "EMPTY!" : formDigestValue![..Math.Min(30, formDigestValue.Length)] + "...")}");
        return formDigestValue;
    }

    /// <summary>
    /// Uses CSOM ProcessQuery to batch-check HasUniqueRoleAssignments for a list of item IDs.
    /// This is more reliable than the REST API for detecting file-level unique permissions.
    /// </summary>
    private async Task<HashSet<int>> CheckUniquePermissionsCsomAsync(
        string baseUrl, string? formDigest, string listTitle, List<int> itemIds)
    {
        var uniqueIds = new HashSet<int>();
        if (itemIds.Count == 0) return uniqueIds;

        Trace.WriteLine($"[SPOManager] CSOM: Checking HasUniqueRoleAssignments for {itemIds.Count} items in '{listTitle}'");

        const int batchSize = 200;
        int batchNum = 0;
        for (int batchStart = 0; batchStart < itemIds.Count; batchStart += batchSize)
        {
            batchNum++;
            var batch = itemIds.Skip(batchStart).Take(batchSize).ToList();

            Trace.WriteLine($"[SPOManager] CSOM: Batch {batchNum} — {batch.Count} items (IDs {batch.First()}..{batch.Last()})");

            var xml = BuildCsomHasUniquePermissionsXml(listTitle, batch);
            var csomUrl = $"{baseUrl}/_vti_bin/client.svc/ProcessQuery";

            Trace.WriteLine($"[SPOManager] CSOM: POST {csomUrl}");
            Trace.WriteLine($"[SPOManager] CSOM: Request XML ({xml.Length} chars): {xml[..Math.Min(500, xml.Length)]}...");

            var request = new HttpRequestMessage(HttpMethod.Post, csomUrl);
            request.Content = new StringContent(xml, System.Text.Encoding.UTF8, "text/xml");
            if (!string.IsNullOrEmpty(formDigest))
                request.Headers.Add("X-RequestDigest", formDigest);

            HttpResponseMessage response;
            try
            {
                response = await _client.SendAsync(request);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SPOManager] CSOM: ProcessQuery EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                continue;
            }

            Trace.WriteLine($"[SPOManager] CSOM: Response HTTP {(int)response.StatusCode} {response.StatusCode}");

            if (!response.IsSuccessStatusCode)
            {
                var errorBody = await response.Content.ReadAsStringAsync();
                Trace.WriteLine($"[SPOManager] CSOM: ERROR body: {errorBody[..Math.Min(1000, errorBody.Length)]}");
                continue;
            }

            var json = await response.Content.ReadAsStringAsync();
            Trace.WriteLine($"[SPOManager] CSOM: Response JSON ({json.Length} chars): {json[..Math.Min(2000, json.Length)]}");

            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                var arrayLen = root.GetArrayLength();

                Trace.WriteLine($"[SPOManager] CSOM: Response array has {arrayLen} elements");

                if (arrayLen == 0)
                {
                    Trace.WriteLine($"[SPOManager] CSOM: Empty response array — skipping batch");
                    continue;
                }

                // Check header for errors
                var header = root[0];
                Trace.WriteLine($"[SPOManager] CSOM: Header: {header}");

                if (header.TryGetProperty("ErrorInfo", out var errorInfo) &&
                    errorInfo.ValueKind != JsonValueKind.Null)
                {
                    var errorMsg = errorInfo.TryGetProperty("ErrorMessage", out var em) ? em.GetString() : "unknown";
                    var errorCode = errorInfo.TryGetProperty("ErrorCode", out var ec) ? ec.GetInt32().ToString() : "?";
                    var errorType = errorInfo.TryGetProperty("ErrorTypeName", out var et) ? et.GetString() : "?";
                    Trace.WriteLine($"[SPOManager] CSOM: BATCH ERROR: [{errorCode}] {errorType}: {errorMsg}");
                    continue;
                }

                // Parse results: array has [header, actionId, resultObj, actionId, resultObj, ...]
                // Our Query action IDs are 2000+batchIndex → maps to batch[batchIndex]
                int uniqueInBatch = 0;
                int queriesFound = 0;
                for (int i = 1; i < arrayLen - 1; i++)
                {
                    var element = root[i];
                    if (element.ValueKind == JsonValueKind.Number)
                    {
                        int actionId = element.GetInt32();

                        if (i + 1 < arrayLen)
                        {
                            i++;
                            var resultObj = root[i];

                            // Only process our Query actions (IDs 2000+)
                            int batchIndex = actionId - 2000;
                            if (batchIndex >= 0 && batchIndex < batch.Count)
                            {
                                queriesFound++;
                                bool hasUnique = resultObj.ValueKind == JsonValueKind.Object &&
                                    resultObj.TryGetProperty("HasUniqueRoleAssignments", out var hasUniqueProp) &&
                                    hasUniqueProp.ValueKind == JsonValueKind.True;

                                if (hasUnique)
                                {
                                    uniqueInBatch++;
                                    uniqueIds.Add(batch[batchIndex]);
                                    Trace.WriteLine($"[SPOManager] CSOM:   Item ID {batch[batchIndex]} → HasUniqueRoleAssignments=TRUE");
                                }
                            }
                        }
                    }
                }

                Trace.WriteLine($"[SPOManager] CSOM: Batch {batchNum} done — {queriesFound} queries parsed, {uniqueInBatch} items with unique perms");
                if (queriesFound != batch.Count)
                    Trace.WriteLine($"[SPOManager] CSOM: WARNING — expected {batch.Count} query results but found {queriesFound}");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[SPOManager] CSOM: PARSE EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                Trace.WriteLine($"[SPOManager] CSOM: Raw JSON: {json[..Math.Min(2000, json.Length)]}");
            }
        }

        Trace.WriteLine($"[SPOManager] CSOM: TOTAL — {uniqueIds.Count} items with unique permissions out of {itemIds.Count}");
        return uniqueIds;
    }

    /// <summary>
    /// Builds CSOM XML to batch-check HasUniqueRoleAssignments for multiple items.
    /// </summary>
    private static string BuildCsomHasUniquePermissionsXml(string listTitle, List<int> itemIds)
    {
        var sb = new System.Text.StringBuilder();
        sb.Append("<Request xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\" SchemaVersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" ApplicationName=\"SPOManager\">");

        // Actions
        sb.Append("<Actions>");
        sb.Append("<ObjectPath Id=\"1\" ObjectPathId=\"10\" />");
        sb.Append("<ObjectPath Id=\"2\" ObjectPathId=\"20\" />");
        sb.Append("<ObjectPath Id=\"3\" ObjectPathId=\"30\" />");
        sb.Append("<ObjectPath Id=\"4\" ObjectPathId=\"40\" />");

        for (int i = 0; i < itemIds.Count; i++)
        {
            int objectPathId = 1000 + i;
            int actionId = 100 + i;
            int queryId = 2000 + i;

            sb.Append($"<ObjectPath Id=\"{actionId}\" ObjectPathId=\"{objectPathId}\" />");
            sb.Append($"<Query Id=\"{queryId}\" ObjectPathId=\"{objectPathId}\">");
            sb.Append("<Query SelectAllProperties=\"false\">");
            sb.Append("<Properties>");
            sb.Append("<Property Name=\"HasUniqueRoleAssignments\" ScalarProperty=\"true\" />");
            sb.Append("</Properties>");
            sb.Append("</Query>");
            sb.Append("</Query>");
        }

        sb.Append("</Actions>");

        // ObjectPaths
        sb.Append("<ObjectPaths>");
        sb.Append("<StaticProperty Id=\"10\" TypeId=\"{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}\" Name=\"Current\" />");
        sb.Append("<Property Id=\"20\" ParentId=\"10\" Name=\"Web\" />");
        sb.Append("<Property Id=\"30\" ParentId=\"20\" Name=\"Lists\" />");

        var escapedTitle = System.Security.SecurityElement.Escape(listTitle);
        sb.Append($"<Method Id=\"40\" ParentId=\"30\" Name=\"GetByTitle\">");
        sb.Append($"<Parameters><Parameter Type=\"String\">{escapedTitle}</Parameter></Parameters>");
        sb.Append("</Method>");

        for (int i = 0; i < itemIds.Count; i++)
        {
            int objectPathId = 1000 + i;
            sb.Append($"<Method Id=\"{objectPathId}\" ParentId=\"40\" Name=\"GetItemById\">");
            sb.Append($"<Parameters><Parameter Type=\"Number\">{itemIds[i]}</Parameter></Parameters>");
            sb.Append("</Method>");
        }

        sb.Append("</ObjectPaths>");
        sb.Append("</Request>");

        return sb.ToString();
    }

    private static SharePointResultStatus GetSharePointStatus(HttpStatusCode code) => code switch
    {
        HttpStatusCode.Unauthorized => SharePointResultStatus.AuthenticationRequired,
        HttpStatusCode.Forbidden => SharePointResultStatus.AccessDenied,
        HttpStatusCode.NotFound => SharePointResultStatus.NotFound,
        _ => SharePointResultStatus.Error
    };

    private static List<PermissionReportItem> ParseRoleAssignments(
        string json,
        string siteCollectionUrl,
        string siteUrl,
        string siteTitle,
        PermissionObjectType objectType,
        string objectTitle,
        string objectUrl,
        bool isInherited,
        string inheritedFrom)
    {
        var items = new List<PermissionReportItem>();

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            JsonElement valueElement;
            if (root.TryGetProperty("value", out valueElement))
            {
                // Standard OData
            }
            else if (root.TryGetProperty("d", out var dElement) &&
                     dElement.TryGetProperty("results", out valueElement))
            {
                // OData verbose
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[SPOManager] ParseRoleAssignments - Could not find 'value' or 'd.results'");
                return items;
            }

            int assignmentCount = 0;
            foreach (var assignment in valueElement.EnumerateArray())
            {
                assignmentCount++;

                // Get member info
                if (!assignment.TryGetProperty("Member", out var member))
                {
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] ParseRoleAssignments - Assignment {assignmentCount}: No 'Member' property");
                    continue;
                }

                var principalName = GetStringProperty(member, "Title");
                var principalLogin = GetStringProperty(member, "LoginName");
                var principalType = GetIntProperty(member, "PrincipalType");

                var principalTypeStr = principalType switch
                {
                    1 => "User",
                    2 => "Distribution List",
                    4 => "Security Group",
                    8 => "SharePoint Group",
                    _ => "Unknown"
                };

                // Get role definition bindings (permission levels)
                if (!assignment.TryGetProperty("RoleDefinitionBindings", out var roleBindings))
                {
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] ParseRoleAssignments - Assignment {assignmentCount}: No 'RoleDefinitionBindings' property");
                    continue;
                }

                // IMPORTANT: Check if it's an array FIRST before trying TryGetProperty
                // With odata=nometadata, RoleDefinitionBindings is a direct array
                // With odata=verbose, it's wrapped in {"results": [...]}
                JsonElement rolesArray;
                if (roleBindings.ValueKind == JsonValueKind.Array)
                {
                    // odata=nometadata format - direct array
                    rolesArray = roleBindings;
                }
                else if (roleBindings.ValueKind == JsonValueKind.Object && roleBindings.TryGetProperty("results", out rolesArray))
                {
                    // OData verbose format - wrapped in results
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] ParseRoleAssignments - Assignment {assignmentCount}: RoleDefinitionBindings is not array (ValueKind={roleBindings.ValueKind})");
                    continue;
                }

                var permissionLevels = new List<string>();
                foreach (var role in rolesArray.EnumerateArray())
                {
                    var roleName = GetStringProperty(role, "Name");
                    if (!string.IsNullOrEmpty(roleName) && !roleName.Equals("Limited Access", StringComparison.OrdinalIgnoreCase))
                    {
                        permissionLevels.Add(roleName);
                    }
                }

                if (permissionLevels.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] ParseRoleAssignments - Assignment {assignmentCount}: {principalName} - Only Limited Access, skipping");
                    continue;
                }

                System.Diagnostics.Debug.WriteLine($"[SPOManager] ParseRoleAssignments - Assignment {assignmentCount}: {principalName} ({principalTypeStr}) = {string.Join(", ", permissionLevels)}");

                items.Add(new PermissionReportItem
                {
                    SiteCollectionUrl = siteCollectionUrl,
                    SiteUrl = siteUrl,
                    SiteTitle = siteTitle,
                    ObjectType = objectType,
                    ObjectTitle = objectTitle,
                    ObjectUrl = objectUrl,
                    PrincipalName = principalName,
                    PrincipalType = principalTypeStr,
                    PrincipalLogin = principalLogin,
                    PermissionLevel = string.Join(", ", permissionLevels),
                    IsInherited = isInherited,
                    InheritedFrom = isInherited ? inheritedFrom : string.Empty
                });
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] ParseRoleAssignments - Processed {assignmentCount} assignments, created {items.Count} entries");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] ParseRoleAssignments - EXCEPTION: {ex.Message}");
        }

        return items;
    }

    #endregion

    #region Site Deletion

    public async Task<SharePointResult<bool>> DeleteSiteCollectionAsync(string siteUrl)
    {
        try
        {
            // Extract tenant name from URL to construct admin URL
            var siteUri = new Uri(siteUrl);
            var hostParts = siteUri.Host.Split('.');
            if (hostParts.Length < 3)
            {
                return new SharePointResult<bool>
                {
                    Status = SharePointResultStatus.Error,
                    ErrorMessage = "Invalid SharePoint URL format"
                };
            }

            var tenantName = hostParts[0];
            var adminUrl = $"https://{tenantName}-admin.sharepoint.com";

            // Get request digest for POST operation
            var digestUrl = $"{adminUrl}/_api/contextinfo";
            var digestResponse = await _client.PostAsync(digestUrl, new StringContent(""));

            if (!digestResponse.IsSuccessStatusCode)
            {
                return new SharePointResult<bool>
                {
                    Status = GetSharePointStatus(digestResponse.StatusCode),
                    ErrorMessage = $"Failed to get request digest: {GetErrorMessage(digestResponse.StatusCode)}"
                };
            }

            var digestJson = await digestResponse.Content.ReadAsStringAsync();
            using var digestDoc = JsonDocument.Parse(digestJson);
            var formDigestValue = GetStringProperty(digestDoc.RootElement, "FormDigestValue");

            if (string.IsNullOrEmpty(formDigestValue))
            {
                // Try alternative path for the digest value
                if (digestDoc.RootElement.TryGetProperty("d", out var dElement) &&
                    dElement.TryGetProperty("GetContextWebInformation", out var webInfo))
                {
                    formDigestValue = GetStringProperty(webInfo, "FormDigestValue");
                }
            }

            // Use Tenant RemoveSite endpoint
            var apiUrl = $"{adminUrl}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/RemoveSite";

            var requestBody = JsonSerializer.Serialize(new { siteUrl = siteUrl });
            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
            {
                Content = new StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
            };

            if (!string.IsNullOrEmpty(formDigestValue))
            {
                request.Headers.Add("X-RequestDigest", formDigestValue);
            }

            var response = await _client.SendAsync(request);

            if (response.IsSuccessStatusCode)
            {
                return new SharePointResult<bool>
                {
                    Data = true,
                    Status = SharePointResultStatus.Success
                };
            }

            return new SharePointResult<bool>
            {
                Data = false,
                Status = GetSharePointStatus(response.StatusCode),
                ErrorMessage = $"Failed to delete site: {GetErrorMessage(response.StatusCode)}"
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<bool>
            {
                Data = false,
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    /// <summary>
    /// Sets the lock state of a site collection.
    /// </summary>
    /// <param name="siteUrl">The URL of the site collection.</param>
    /// <param name="lockState">The lock state: "Unlock", "ReadOnly", or "NoAccess".</param>
    public async Task<SharePointResult<bool>> SetSiteLockStateAsync(string siteUrl, string lockState)
    {
        try
        {
            // Extract tenant name from URL to construct admin URL
            var siteUri = new Uri(siteUrl);
            var hostParts = siteUri.Host.Split('.');
            if (hostParts.Length < 3)
            {
                return new SharePointResult<bool>
                {
                    Status = SharePointResultStatus.Error,
                    ErrorMessage = "Invalid SharePoint URL format"
                };
            }

            var tenantName = hostParts[0];
            var adminUrl = $"https://{tenantName}-admin.sharepoint.com";
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - SiteUrl: {siteUrl}, LockState: {lockState}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - Admin URL: {adminUrl}, Cookie Domain: {_domain}");

            // Get request digest for POST operation
            var digestUrl = $"{adminUrl}/_api/contextinfo";
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - Getting digest from: {digestUrl}");
            var digestResponse = await _client.PostAsync(digestUrl, new StringContent(""));
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - Digest response: {digestResponse.StatusCode}");

            if (!digestResponse.IsSuccessStatusCode)
            {
                var digestError = await digestResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - Digest error: {digestError}");
                return new SharePointResult<bool>
                {
                    Status = GetSharePointStatus(digestResponse.StatusCode),
                    ErrorMessage = $"Failed to get request digest: {digestResponse.StatusCode} - {digestError}"
                };
            }

            var digestJson = await digestResponse.Content.ReadAsStringAsync();
            using var digestDoc = JsonDocument.Parse(digestJson);
            var formDigestValue = GetStringProperty(digestDoc.RootElement, "FormDigestValue");

            if (string.IsNullOrEmpty(formDigestValue))
            {
                // Try alternative path for the digest value
                if (digestDoc.RootElement.TryGetProperty("d", out var dElement) &&
                    dElement.TryGetProperty("GetContextWebInformation", out var webInfo))
                {
                    formDigestValue = GetStringProperty(webInfo, "FormDigestValue");
                }
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - Got digest: {!string.IsNullOrEmpty(formDigestValue)}");

            // Use CSOM ProcessQuery endpoint - this is how SharePoint Admin operations work
            var apiUrl = $"{adminUrl}/_vti_bin/client.svc/ProcessQuery";

            // CSOM XML request to set site lock state
            // LockState valid values are: "Unlock", "ReadOnly", "NoAccess"
            var csomRequest = $@"<Request xmlns=""http://schemas.microsoft.com/sharepoint/clientquery/2009"" SchemaVersion=""15.0.0.0"" LibraryVersion=""16.0.0.0"" ApplicationName=""SharePoint Online Manager"">
  <Actions>
    <ObjectPath Id=""1"" ObjectPathId=""0"" />
    <ObjectPath Id=""3"" ObjectPathId=""2"" />
    <SetProperty Id=""4"" ObjectPathId=""2"" Name=""LockState"">
      <Parameter Type=""String"">{lockState}</Parameter>
    </SetProperty>
    <Method Name=""Update"" Id=""5"" ObjectPathId=""2"" />
  </Actions>
  <ObjectPaths>
    <Constructor Id=""0"" TypeId=""{{268004ae-ef6b-4e9b-8425-127220d84719}}"" />
    <Method Id=""2"" ParentId=""0"" Name=""GetSitePropertiesByUrl"">
      <Parameters>
        <Parameter Type=""String"">{System.Security.SecurityElement.Escape(siteUrl)}</Parameter>
        <Parameter Type=""Boolean"">false</Parameter>
      </Parameters>
    </Method>
  </ObjectPaths>
</Request>";

            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
            {
                Content = new StringContent(csomRequest, System.Text.Encoding.UTF8, "text/xml")
            };

            if (!string.IsNullOrEmpty(formDigestValue))
            {
                request.Headers.Add("X-RequestDigest", formDigestValue);
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - URL: {apiUrl}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - CSOM Request sent");

            var response = await _client.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - Response: {response.StatusCode}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - Response content: {responseContent}");

            if (response.IsSuccessStatusCode)
            {
                // Check if the CSOM response contains an error
                if (responseContent.Contains("ErrorInfo") && responseContent.Contains("ErrorMessage"))
                {
                    // Extract error message from CSOM response
                    var errorStart = responseContent.IndexOf("ErrorMessage");
                    var errorEnd = responseContent.IndexOf(",", errorStart);
                    var errorMsg = errorStart > 0 && errorEnd > errorStart
                        ? responseContent.Substring(errorStart, errorEnd - errorStart)
                        : "CSOM operation failed";

                    return new SharePointResult<bool>
                    {
                        Data = false,
                        Status = SharePointResultStatus.Error,
                        ErrorMessage = errorMsg
                    };
                }

                return new SharePointResult<bool>
                {
                    Data = true,
                    Status = SharePointResultStatus.Success
                };
            }

            return new SharePointResult<bool>
            {
                Data = false,
                Status = GetSharePointStatus(response.StatusCode),
                ErrorMessage = $"{response.StatusCode}: {responseContent}"
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetSiteLockState - Exception: {ex.Message}");
            return new SharePointResult<bool>
            {
                Data = false,
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    #endregion

    #region User Search and Site Admin Methods

    public async Task<SharePointResult<List<UserSearchResult>>> SearchUsersAsync(string siteUrl, string queryString)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(queryString) || queryString.Length < 2)
            {
                return new SharePointResult<List<UserSearchResult>>
                {
                    Data = [],
                    Status = SharePointResultStatus.Success
                };
            }

            var baseUrl = siteUrl.TrimEnd('/');
            var apiUrl = $"{baseUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser";
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SearchUsersAsync - URL: {apiUrl}, Query: {queryString}");

            // Get request digest first
            var digestUrl = $"{baseUrl}/_api/contextinfo";
            var digestResponse = await _client.PostAsync(digestUrl, new StringContent(""));
            string? formDigestValue = null;

            if (digestResponse.IsSuccessStatusCode)
            {
                var digestJson = await digestResponse.Content.ReadAsStringAsync();
                using var digestDoc = JsonDocument.Parse(digestJson);
                formDigestValue = GetStringProperty(digestDoc.RootElement, "FormDigestValue");

                if (string.IsNullOrEmpty(formDigestValue))
                {
                    if (digestDoc.RootElement.TryGetProperty("d", out var dElement) &&
                        dElement.TryGetProperty("GetContextWebInformation", out var webInfo))
                    {
                        formDigestValue = GetStringProperty(webInfo, "FormDigestValue");
                    }
                }
                System.Diagnostics.Debug.WriteLine($"[SPOManager] SearchUsersAsync - Got digest: {!string.IsNullOrEmpty(formDigestValue)}");
            }

            // Build the people picker query
            var queryParams = new
            {
                queryParams = new
                {
                    AllowEmailAddresses = true,
                    AllowMultipleEntities = false,
                    AllUrlZones = false,
                    MaximumEntitySuggestions = 10,
                    PrincipalSource = 15, // All sources
                    PrincipalType = 1, // Users only (1=User, 4=SecurityGroup, 8=SharePointGroup)
                    QueryString = queryString
                }
            };

            var requestBody = JsonSerializer.Serialize(queryParams);
            var content = new StringContent(requestBody, System.Text.Encoding.UTF8, "application/json");

            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
            {
                Content = content
            };

            if (!string.IsNullOrEmpty(formDigestValue))
            {
                request.Headers.Add("X-RequestDigest", formDigestValue);
            }

            var response = await _client.SendAsync(request);
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SearchUsersAsync - Response status: {response.StatusCode}");

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager] SearchUsersAsync - Error: {errorContent}");
                return new SharePointResult<List<UserSearchResult>>
                {
                    Status = GetSharePointStatus(response.StatusCode),
                    ErrorMessage = GetErrorMessage(response.StatusCode)
                };
            }

            var json = await response.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SearchUsersAsync - Response: {json.Substring(0, Math.Min(500, json.Length))}");
            var users = ParsePeoplePickerResults(json);

            return new SharePointResult<List<UserSearchResult>>
            {
                Data = users,
                Status = SharePointResultStatus.Success
            };
        }
        catch (HttpRequestException ex)
        {
            return new SharePointResult<List<UserSearchResult>>
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
            return new SharePointResult<List<UserSearchResult>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    private static List<UserSearchResult> ParsePeoplePickerResults(string json)
    {
        var users = new List<UserSearchResult>();

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            // The result is in "value" property as a JSON string that needs to be parsed again
            string? valueString = null;
            if (root.TryGetProperty("value", out var valueElement))
            {
                valueString = valueElement.GetString();
            }
            else if (root.TryGetProperty("d", out var dElement) &&
                     dElement.TryGetProperty("ClientPeoplePickerSearchUser", out var searchElement))
            {
                valueString = searchElement.GetString();
            }

            if (string.IsNullOrEmpty(valueString))
                return users;

            // Parse the nested JSON array
            using var resultsDoc = JsonDocument.Parse(valueString);
            var resultsArray = resultsDoc.RootElement;

            foreach (var item in resultsArray.EnumerateArray())
            {
                var displayName = GetStringProperty(item, "DisplayText");
                var entityType = GetStringProperty(item, "EntityType");

                // Get entity data for email and login name
                var email = string.Empty;
                var loginName = string.Empty;

                if (item.TryGetProperty("EntityData", out var entityData))
                {
                    email = GetStringProperty(entityData, "Email");
                }

                if (item.TryGetProperty("Key", out var keyProp))
                {
                    loginName = keyProp.GetString() ?? string.Empty;
                }

                if (!string.IsNullOrEmpty(displayName))
                {
                    users.Add(new UserSearchResult
                    {
                        DisplayName = displayName,
                        Email = email,
                        LoginName = loginName,
                        EntityType = entityType
                    });
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] ParsePeoplePickerResults - Exception: {ex.Message}");
        }

        return users;
    }

    public async Task<SharePointResult<bool>> AddSiteCollectionAdminAsync(string siteUrl, string userLoginName)
    {
        try
        {
            // Extract tenant name from URL to construct admin URL
            var siteUri = new Uri(siteUrl);
            var hostParts = siteUri.Host.Split('.');
            if (hostParts.Length < 3)
            {
                return new SharePointResult<bool>
                {
                    Status = SharePointResultStatus.Error,
                    ErrorMessage = "Invalid SharePoint URL format"
                };
            }

            var tenantName = hostParts[0];
            var adminUrl = $"https://{tenantName}-admin.sharepoint.com";
            System.Diagnostics.Debug.WriteLine($"[SPOManager] AddSiteCollectionAdmin - SiteUrl: {siteUrl}, User: {userLoginName}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] AddSiteCollectionAdmin - Admin URL: {adminUrl}");

            // Get request digest for POST operation
            var digestUrl = $"{adminUrl}/_api/contextinfo";
            var digestResponse = await _client.PostAsync(digestUrl, new StringContent(""));

            if (!digestResponse.IsSuccessStatusCode)
            {
                var digestError = await digestResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"[SPOManager] AddSiteCollectionAdmin - Digest error: {digestError}");
                return new SharePointResult<bool>
                {
                    Status = GetSharePointStatus(digestResponse.StatusCode),
                    ErrorMessage = $"Failed to get request digest: {digestResponse.StatusCode}"
                };
            }

            var digestJson = await digestResponse.Content.ReadAsStringAsync();
            using var digestDoc = JsonDocument.Parse(digestJson);
            var formDigestValue = GetStringProperty(digestDoc.RootElement, "FormDigestValue");

            if (string.IsNullOrEmpty(formDigestValue))
            {
                if (digestDoc.RootElement.TryGetProperty("d", out var dElement) &&
                    dElement.TryGetProperty("GetContextWebInformation", out var webInfo))
                {
                    formDigestValue = GetStringProperty(webInfo, "FormDigestValue");
                }
            }

            // Use CSOM ProcessQuery endpoint with Tenant.SetSiteAdmin
            var apiUrl = $"{adminUrl}/_vti_bin/client.svc/ProcessQuery";

            // CSOM XML request to add site collection admin
            var csomRequest = $@"<Request xmlns=""http://schemas.microsoft.com/sharepoint/clientquery/2009"" SchemaVersion=""15.0.0.0"" LibraryVersion=""16.0.0.0"" ApplicationName=""SharePoint Online Manager"">
  <Actions>
    <ObjectPath Id=""1"" ObjectPathId=""0"" />
    <Method Name=""SetSiteAdmin"" Id=""2"" ObjectPathId=""0"">
      <Parameters>
        <Parameter Type=""String"">{System.Security.SecurityElement.Escape(siteUrl)}</Parameter>
        <Parameter Type=""String"">{System.Security.SecurityElement.Escape(userLoginName)}</Parameter>
        <Parameter Type=""Boolean"">true</Parameter>
      </Parameters>
    </Method>
  </Actions>
  <ObjectPaths>
    <Constructor Id=""0"" TypeId=""{{268004ae-ef6b-4e9b-8425-127220d84719}}"" />
  </ObjectPaths>
</Request>";

            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
            {
                Content = new StringContent(csomRequest, System.Text.Encoding.UTF8, "text/xml")
            };

            if (!string.IsNullOrEmpty(formDigestValue))
            {
                request.Headers.Add("X-RequestDigest", formDigestValue);
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] AddSiteCollectionAdmin - Sending CSOM request");

            var response = await _client.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            System.Diagnostics.Debug.WriteLine($"[SPOManager] AddSiteCollectionAdmin - Response: {response.StatusCode}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] AddSiteCollectionAdmin - Content: {responseContent}");

            if (response.IsSuccessStatusCode)
            {
                // Check if the CSOM response contains an error
                if (responseContent.Contains("ErrorInfo") && responseContent.Contains("ErrorMessage"))
                {
                    // Extract error message from CSOM response
                    var errorStart = responseContent.IndexOf("ErrorMessage");
                    var errorEnd = responseContent.IndexOf(",", errorStart);
                    var errorMsg = errorStart > 0 && errorEnd > errorStart
                        ? responseContent.Substring(errorStart, errorEnd - errorStart)
                        : "CSOM operation failed";

                    return new SharePointResult<bool>
                    {
                        Data = false,
                        Status = SharePointResultStatus.Error,
                        ErrorMessage = errorMsg
                    };
                }

                return new SharePointResult<bool>
                {
                    Data = true,
                    Status = SharePointResultStatus.Success
                };
            }

            return new SharePointResult<bool>
            {
                Data = false,
                Status = GetSharePointStatus(response.StatusCode),
                ErrorMessage = $"{response.StatusCode}: {responseContent}"
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] AddSiteCollectionAdmin - Exception: {ex.Message}");
            return new SharePointResult<bool>
            {
                Data = false,
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    public async Task<SharePointResult<bool>> RemoveSiteCollectionAdminAsync(string siteUrl, string userLoginName)
    {
        try
        {
            var siteUri = new Uri(siteUrl);
            var hostParts = siteUri.Host.Split('.');
            if (hostParts.Length < 3)
            {
                return new SharePointResult<bool>
                {
                    Status = SharePointResultStatus.Error,
                    ErrorMessage = "Invalid SharePoint URL format"
                };
            }

            var tenantName = hostParts[0];
            var adminUrl = $"https://{tenantName}-admin.sharepoint.com";
            System.Diagnostics.Debug.WriteLine($"[SPOManager] RemoveSiteCollectionAdmin - SiteUrl: {siteUrl}, User: {userLoginName}");

            var digestUrl = $"{adminUrl}/_api/contextinfo";
            var digestResponse = await _client.PostAsync(digestUrl, new StringContent(""));

            if (!digestResponse.IsSuccessStatusCode)
            {
                return new SharePointResult<bool>
                {
                    Status = GetSharePointStatus(digestResponse.StatusCode),
                    ErrorMessage = $"Failed to get request digest: {digestResponse.StatusCode}"
                };
            }

            var digestJson = await digestResponse.Content.ReadAsStringAsync();
            using var digestDoc = JsonDocument.Parse(digestJson);
            var formDigestValue = GetStringProperty(digestDoc.RootElement, "FormDigestValue");

            if (string.IsNullOrEmpty(formDigestValue))
            {
                if (digestDoc.RootElement.TryGetProperty("d", out var dElement) &&
                    dElement.TryGetProperty("GetContextWebInformation", out var webInfo))
                {
                    formDigestValue = GetStringProperty(webInfo, "FormDigestValue");
                }
            }

            var apiUrl = $"{adminUrl}/_vti_bin/client.svc/ProcessQuery";

            // CSOM XML request to remove site collection admin (false instead of true)
            var csomRequest = $@"<Request xmlns=""http://schemas.microsoft.com/sharepoint/clientquery/2009"" SchemaVersion=""15.0.0.0"" LibraryVersion=""16.0.0.0"" ApplicationName=""SharePoint Online Manager"">
  <Actions>
    <ObjectPath Id=""1"" ObjectPathId=""0"" />
    <Method Name=""SetSiteAdmin"" Id=""2"" ObjectPathId=""0"">
      <Parameters>
        <Parameter Type=""String"">{System.Security.SecurityElement.Escape(siteUrl)}</Parameter>
        <Parameter Type=""String"">{System.Security.SecurityElement.Escape(userLoginName)}</Parameter>
        <Parameter Type=""Boolean"">false</Parameter>
      </Parameters>
    </Method>
  </Actions>
  <ObjectPaths>
    <Constructor Id=""0"" TypeId=""{{268004ae-ef6b-4e9b-8425-127220d84719}}"" />
  </ObjectPaths>
</Request>";

            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
            {
                Content = new StringContent(csomRequest, System.Text.Encoding.UTF8, "text/xml")
            };

            if (!string.IsNullOrEmpty(formDigestValue))
            {
                request.Headers.Add("X-RequestDigest", formDigestValue);
            }

            var response = await _client.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            System.Diagnostics.Debug.WriteLine($"[SPOManager] RemoveSiteCollectionAdmin - Response: {response.StatusCode}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] RemoveSiteCollectionAdmin - Content: {responseContent}");

            if (response.IsSuccessStatusCode)
            {
                if (responseContent.Contains("ErrorInfo") && responseContent.Contains("ErrorMessage"))
                {
                    var errorStart = responseContent.IndexOf("ErrorMessage");
                    var errorEnd = responseContent.IndexOf(",", errorStart);
                    var errorMsg = errorStart > 0 && errorEnd > errorStart
                        ? responseContent.Substring(errorStart, errorEnd - errorStart)
                        : "CSOM operation failed";

                    return new SharePointResult<bool>
                    {
                        Data = false,
                        Status = SharePointResultStatus.Error,
                        ErrorMessage = errorMsg
                    };
                }

                return new SharePointResult<bool>
                {
                    Data = true,
                    Status = SharePointResultStatus.Success
                };
            }

            return new SharePointResult<bool>
            {
                Data = false,
                Status = GetSharePointStatus(response.StatusCode),
                ErrorMessage = $"{response.StatusCode}: {responseContent}"
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] RemoveSiteCollectionAdmin - Exception: {ex.Message}");
            return new SharePointResult<bool>
            {
                Data = false,
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    #endregion

    #region Navigation Settings Methods

    public async Task<SharePointResult<NavigationSettings>> GetNavigationSettingsAsync(string siteUrl)
    {
        try
        {
            var baseUrl = siteUrl.TrimEnd('/');
            var apiUrl = $"{baseUrl}/_api/web?$select=HorizontalQuickLaunch,MegaMenuEnabled";
            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetNavigationSettingsAsync - URL: {apiUrl}");

            var response = await _client.GetAsync(apiUrl);

            if (!response.IsSuccessStatusCode)
            {
                return new SharePointResult<NavigationSettings>
                {
                    Status = GetSharePointStatus(response.StatusCode),
                    ErrorMessage = GetErrorMessage(response.StatusCode)
                };
            }

            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            // Handle both formats
            var webData = root;
            if (webData.TryGetProperty("d", out var dElement))
            {
                webData = dElement;
            }

            var settings = new NavigationSettings
            {
                HorizontalQuickLaunch = GetBoolProperty(webData, "HorizontalQuickLaunch"),
                MegaMenuEnabled = GetBoolProperty(webData, "MegaMenuEnabled")
            };

            System.Diagnostics.Debug.WriteLine($"[SPOManager] GetNavigationSettingsAsync - HorizontalQuickLaunch: {settings.HorizontalQuickLaunch}, MegaMenuEnabled: {settings.MegaMenuEnabled}");

            return new SharePointResult<NavigationSettings>
            {
                Data = settings,
                Status = SharePointResultStatus.Success
            };
        }
        catch (HttpRequestException ex)
        {
            return new SharePointResult<NavigationSettings>
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
            return new SharePointResult<NavigationSettings>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    public async Task<SharePointResult<bool>> SetNavigationSettingsAsync(string siteUrl, NavigationSettings settings)
    {
        try
        {
            var baseUrl = siteUrl.TrimEnd('/');
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetNavigationSettingsAsync - SiteUrl: {siteUrl}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   HorizontalQuickLaunch: {settings.HorizontalQuickLaunch}, MegaMenuEnabled: {settings.MegaMenuEnabled}");

            // Get request digest first
            var digestUrl = $"{baseUrl}/_api/contextinfo";
            var digestResponse = await _client.PostAsync(digestUrl, new StringContent(""));

            if (!digestResponse.IsSuccessStatusCode)
            {
                return new SharePointResult<bool>
                {
                    Status = GetSharePointStatus(digestResponse.StatusCode),
                    ErrorMessage = $"Failed to get request digest: {GetErrorMessage(digestResponse.StatusCode)}"
                };
            }

            var digestJson = await digestResponse.Content.ReadAsStringAsync();
            using var digestDoc = JsonDocument.Parse(digestJson);
            var formDigestValue = GetStringProperty(digestDoc.RootElement, "FormDigestValue");

            if (string.IsNullOrEmpty(formDigestValue))
            {
                if (digestDoc.RootElement.TryGetProperty("d", out var dElement) &&
                    dElement.TryGetProperty("GetContextWebInformation", out var webInfo))
                {
                    formDigestValue = GetStringProperty(webInfo, "FormDigestValue");
                }
            }

            // Update the web properties using MERGE
            var apiUrl = $"{baseUrl}/_api/web";
            var updateBody = JsonSerializer.Serialize(new
            {
                HorizontalQuickLaunch = settings.HorizontalQuickLaunch,
                MegaMenuEnabled = settings.MegaMenuEnabled
            });

            var content = new StringContent(updateBody, System.Text.Encoding.UTF8);
            content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json")
            {
                Parameters = { new System.Net.Http.Headers.NameValueHeaderValue("odata", "nometadata") }
            };

            var request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
            {
                Content = content
            };

            request.Headers.Add("X-HTTP-Method", "MERGE");
            request.Headers.Add("IF-MATCH", "*");
            if (!string.IsNullOrEmpty(formDigestValue))
            {
                request.Headers.Add("X-RequestDigest", formDigestValue);
            }

            var response = await _client.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetNavigationSettingsAsync - Response: {response.StatusCode}");

            if (response.IsSuccessStatusCode || response.StatusCode == HttpStatusCode.NoContent)
            {
                return new SharePointResult<bool>
                {
                    Data = true,
                    Status = SharePointResultStatus.Success
                };
            }

            return new SharePointResult<bool>
            {
                Data = false,
                Status = GetSharePointStatus(response.StatusCode),
                ErrorMessage = $"{response.StatusCode}: {responseContent}"
            };
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] SetNavigationSettingsAsync - Exception: {ex.Message}");
            return new SharePointResult<bool>
            {
                Data = false,
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
    }

    #endregion

    #region Document Compare Methods

    /// <summary>
    /// Exponential backoff delay values in seconds for throttling retry.
    /// </summary>
    private static readonly int[] ThrottleRetryDelays = [2, 4, 8, 16, 32, 60, 120];

    /// <summary>
    /// Returns true if the status code indicates SharePoint throttling.
    /// SharePoint Online uses 429 (Too Many Requests) and 503 (Service Unavailable) for throttling.
    /// </summary>
    private static bool IsThrottlingResponse(HttpStatusCode statusCode) =>
        statusCode == (HttpStatusCode)429 || statusCode == HttpStatusCode.ServiceUnavailable;

    /// <summary>
    /// Extracts the Retry-After delay in seconds from the response headers.
    /// SharePoint sends Retry-After as an integer (seconds).
    /// Returns null if no valid Retry-After header is present.
    /// </summary>
    private static int? GetRetryAfterSeconds(HttpResponseMessage response)
    {
        // Try the typed Retry-After header (handles both delta and date formats)
        if (response.Headers.RetryAfter?.Delta.HasValue == true)
        {
            return Math.Max(1, (int)response.Headers.RetryAfter.Delta.Value.TotalSeconds);
        }
        if (response.Headers.RetryAfter?.Date.HasValue == true)
        {
            var waitTime = response.Headers.RetryAfter.Date.Value - DateTimeOffset.UtcNow;
            return Math.Max(1, (int)waitTime.TotalSeconds);
        }

        // Fallback: parse raw header value as integer seconds
        // (some SharePoint responses send it as a plain integer string)
        if (response.Headers.TryGetValues("Retry-After", out var values))
        {
            var raw = values.FirstOrDefault();
            if (!string.IsNullOrEmpty(raw) && int.TryParse(raw, out var seconds))
            {
                return Math.Max(1, seconds);
            }
        }

        return null;
    }

    /// <summary>
    /// Executes an HTTP request with automatic retry for throttling (429/503) responses.
    /// Honors the Retry-After header when present, falls back to exponential backoff.
    /// </summary>
    private async Task<HttpResponseMessage> ExecuteWithThrottlingRetryAsync(
        Func<Task<HttpResponseMessage>> requestFunc,
        CancellationToken cancellationToken)
    {
        int retryCount = 0;

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var response = await requestFunc();

            if (!IsThrottlingResponse(response.StatusCode))
            {
                return response;
            }

            if (retryCount >= ThrottleRetryDelays.Length)
            {
                // Max retries reached, return the throttled response
                return response;
            }

            // Honor Retry-After header if present, otherwise use exponential backoff
            var retryAfterSeconds = GetRetryAfterSeconds(response) ?? ThrottleRetryDelays[retryCount];

            Trace.WriteLine($"[SPOManager] Throttled ({(int)response.StatusCode}) - Retry {retryCount + 1}/{ThrottleRetryDelays.Length}, waiting {retryAfterSeconds}s");

            await Task.Delay(TimeSpan.FromSeconds(retryAfterSeconds), cancellationToken);
            retryCount++;
        }
    }

    public async Task<SharePointResult<List<DocumentCompareSourceItem>>> GetDocumentsForCompareAsync(
        string siteUrl,
        string libraryTitle,
        bool includeAspxPages,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        var documents = new List<DocumentCompareSourceItem>();
        var baseUrl = siteUrl.TrimEnd('/');
        var encodedLibraryTitle = Uri.EscapeDataString(libraryTitle);

        try
        {
            // Query both files (FSObjType=0) and folders (FSObjType=1)
            // ASPX filtering is done during parsing if needed
            var viewXml = @"<View Scope='RecursiveAll'>
                <ViewFields>
                    <FieldRef Name='ID'/>
                    <FieldRef Name='FileLeafRef'/>
                    <FieldRef Name='FileRef'/>
                    <FieldRef Name='File_x0020_Size'/>
                    <FieldRef Name='_UIVersionString'/>
                    <FieldRef Name='FSObjType'/>
                    <FieldRef Name='Created'/>
                    <FieldRef Name='Modified'/>
                </ViewFields>
                <RowLimit Paged='TRUE'>5000</RowLimit>
            </View>";

            string? pagingInfo = null;
            int pageCount = 0;
            var baseApiUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedLibraryTitle}')/RenderListDataAsStream";

            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                pageCount++;

                var fetchedSoFar = documents.Count > 0 ? $" ({documents.Count:N0} fetched)" : "";
                progress?.Report($"Fetching page {pageCount} from {libraryTitle}{fetchedSoFar}...");

                var apiUrl = baseApiUrl;
                if (!string.IsNullOrEmpty(pagingInfo))
                {
                    apiUrl = baseApiUrl + pagingInfo;
                }

                var requestBody = new StringContent(
                    $"{{\"parameters\":{{\"RenderOptions\":2,\"ViewXml\":\"{viewXml.Replace("\"", "\\\"").Replace("\n", "").Replace("\r", "")}\"}}}}",
                    System.Text.Encoding.UTF8,
                    "application/json");

                var response = await ExecuteWithThrottlingRetryAsync(
                    () => _client.PostAsync(apiUrl, requestBody),
                    cancellationToken);

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    return new SharePointResult<List<DocumentCompareSourceItem>>
                    {
                        Data = documents,
                        Status = GetSharePointStatus(response.StatusCode),
                        ErrorMessage = $"Page {pageCount}: HTTP {(int)response.StatusCode} - {errorContent.Substring(0, Math.Min(200, errorContent.Length))}"
                    };
                }

                var json = await response.Content.ReadAsStringAsync();
                var (items, nextPaging) = ParseDocumentsForCompare(json, libraryTitle, includeAspxPages, siteUrl);
                documents.AddRange(items);
                pagingInfo = nextPaging;

            } while (!string.IsNullOrEmpty(pagingInfo));

            return new SharePointResult<List<DocumentCompareSourceItem>>
            {
                Data = documents,
                Status = SharePointResultStatus.Success
            };
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (HttpRequestException ex)
        {
            return new SharePointResult<List<DocumentCompareSourceItem>>
            {
                Data = documents,
                Status = ex.StatusCode switch
                {
                    HttpStatusCode.Unauthorized => SharePointResultStatus.AuthenticationRequired,
                    HttpStatusCode.Forbidden => SharePointResultStatus.AccessDenied,
                    HttpStatusCode.NotFound => SharePointResultStatus.NotFound,
                    _ => SharePointResultStatus.Error
                },
                ErrorMessage = $"{ex.Message}; Collected {documents.Count} docs before error"
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<List<DocumentCompareSourceItem>>
            {
                Data = documents,
                Status = SharePointResultStatus.Error,
                ErrorMessage = $"{ex.Message}; Collected {documents.Count} docs before error"
            };
        }
    }

    private static (List<DocumentCompareSourceItem> items, string? nextHref) ParseDocumentsForCompare(
        string json,
        string libraryTitle,
        bool includeAspxPages,
        string siteUrl)
    {
        var items = new List<DocumentCompareSourceItem>();
        string? nextHref = null;

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            // Get paging info
            if (root.TryGetProperty("NextHref", out var nextHrefElement) &&
                nextHrefElement.ValueKind == JsonValueKind.String)
            {
                nextHref = nextHrefElement.GetString();
            }

            // Get rows
            if (!root.TryGetProperty("Row", out var rowsElement))
            {
                return (items, nextHref);
            }

            foreach (var row in rowsElement.EnumerateArray())
            {
                var fileName = GetStringProperty(row, "FileLeafRef");
                var fileRef = GetStringProperty(row, "FileRef");

                if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(fileRef))
                    continue;

                // Parse FSObjType: 0 = File, 1 = Folder
                var itemType = DocumentCompareItemType.File;
                var fsObjTypeStr = GetStringProperty(row, "FSObjType");
                if (fsObjTypeStr == "1")
                {
                    itemType = DocumentCompareItemType.Folder;
                }

                // Skip ASPX files if not including them (folders are always included)
                if (itemType == DocumentCompareItemType.File && !includeAspxPages)
                {
                    if (fileName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }

                // Parse ID
                int id = 0;
                if (row.TryGetProperty("ID", out var idProp))
                {
                    if (idProp.ValueKind == JsonValueKind.Number)
                    {
                        id = idProp.GetInt32();
                    }
                    else if (idProp.ValueKind == JsonValueKind.String)
                    {
                        int.TryParse(idProp.GetString(), out id);
                    }
                }

                // Parse file size (folders will have 0)
                long fileSize = 0;
                var fileSizeStr = GetStringProperty(row, "File_x0020_Size");
                if (!string.IsNullOrEmpty(fileSizeStr))
                {
                    var cleanSize = fileSizeStr.Replace(",", "").Replace(" ", "");
                    long.TryParse(cleanSize, out fileSize);
                }

                // Parse version (folders typically have version 1)
                var versionStr = GetStringProperty(row, "_UIVersionString");
                int versionCount = 1;
                if (!string.IsNullOrEmpty(versionStr))
                {
                    var dotIndex = versionStr.IndexOf('.');
                    if (dotIndex > 0 && int.TryParse(versionStr.Substring(0, dotIndex), out var major))
                    {
                        versionCount = major;
                    }
                }

                // Parse Created and Modified dates
                // RenderListDataAsStream may return these with a period suffix (e.g., "Created." not "Created")
                DateTime? created = null;
                DateTime? modified = null;
                var createdStr = GetStringProperty(row, "Created.");
                if (string.IsNullOrEmpty(createdStr))
                    createdStr = GetStringProperty(row, "Created");
                var modifiedStr = GetStringProperty(row, "Modified.");
                if (string.IsNullOrEmpty(modifiedStr))
                    modifiedStr = GetStringProperty(row, "Modified");
                if (!string.IsNullOrEmpty(createdStr) && DateTime.TryParse(createdStr, out var createdDt))
                {
                    created = createdDt;
                }
                if (!string.IsNullOrEmpty(modifiedStr) && DateTime.TryParse(modifiedStr, out var modifiedDt))
                {
                    modified = modifiedDt;
                }

                // Build relative path (case-insensitive matching key)
                // Extract the path after the library name for comparison
                var relativePath = ExtractRelativePathForCompare(fileRef, libraryTitle, siteUrl);

                items.Add(new DocumentCompareSourceItem
                {
                    Id = id,
                    FileName = fileName,
                    ServerRelativeUrl = fileRef,
                    RelativePath = relativePath,
                    SizeBytes = fileSize,
                    VersionCount = versionCount,
                    LibraryTitle = libraryTitle,
                    ItemType = itemType,
                    Created = created,
                    Modified = modified
                });
            }
        }
        catch
        {
            // Return what we have
        }

        return (items, nextHref);
    }

    /// <summary>
    /// Re-computes RelativePath for cached documents using current extraction logic.
    /// This ensures cached data stays correct even if the path extraction algorithm changes.
    /// </summary>
    public static void RecomputeRelativePaths(List<DocumentCompareSourceItem> documents, string libraryTitle, string siteUrl)
    {
        foreach (var doc in documents)
        {
            doc.RelativePath = ExtractRelativePathForCompare(doc.ServerRelativeUrl, libraryTitle, siteUrl);
        }
    }

    private static string ExtractRelativePathForCompare(string serverRelativeUrl, string libraryTitle, string siteUrl)
    {
        // ServerRelativeUrl format: /sites/sitename/libraryurl/folder1/folder2/filename.ext
        // Or for root sites: /libraryurl/folder1/folder2/filename.ext
        // We want to extract: folder1/folder2/filename.ext (normalized for comparison)
        // NOTE: Do NOT URL-decode the path. RenderListDataAsStream returns raw field values,
        // and filenames may contain literal %XX sequences (e.g., "file%2E1.pdf") that would
        // be incorrectly decoded to dots/spaces.
        try
        {
            var lowerUrl = serverRelativeUrl.ToLowerInvariant();

            // Strategy 1: Use the site URL to strip the site prefix, then skip the library URL segment.
            // This works even when the library display title differs from the URL name
            // (e.g., "S&T" display title but "ST" URL name, or "UOHI Polices" title but "Polices" URL)
            var siteServerRelativePath = GetServerRelativePath(siteUrl);
            if (!string.IsNullOrEmpty(siteServerRelativePath))
            {
                var lowerSitePath = siteServerRelativePath.ToLowerInvariant().TrimEnd('/') + "/";
                if (lowerUrl.StartsWith(lowerSitePath))
                {
                    // After stripping site path, we have: libraryurl/folder1/folder2/file.ext
                    var afterSite = serverRelativeUrl.Substring(lowerSitePath.Length);
                    var firstSlash = afterSite.IndexOf('/');
                    if (firstSlash >= 0 && firstSlash < afterSite.Length - 1)
                    {
                        // Everything after libraryurl/ is the content path
                        return afterSite.Substring(firstSlash + 1).ToLowerInvariant();
                    }
                    // File is directly in the library root
                    return string.Empty;
                }
            }

            // Strategy 2: Try matching library title variations in the URL
            // Handles cases where site URL extraction didn't work
            var candidates = new List<string>
            {
                libraryTitle.ToLowerInvariant() + "/"
            };

            // Also try without spaces (e.g., "Site Assets" → "SiteAssets")
            var noSpaces = libraryTitle.Replace(" ", "").ToLowerInvariant() + "/";
            if (noSpaces != candidates[0])
            {
                candidates.Add(noSpaces);
            }

            foreach (var candidate in candidates)
            {
                var libraryIndex = lowerUrl.LastIndexOf(candidate, StringComparison.OrdinalIgnoreCase);
                if (libraryIndex >= 0)
                {
                    if (libraryIndex == 0 || serverRelativeUrl[libraryIndex - 1] == '/' || serverRelativeUrl[libraryIndex - 1] == ' ')
                    {
                        var pathStart = libraryIndex + candidate.Length;
                        if (pathStart < serverRelativeUrl.Length)
                        {
                            return serverRelativeUrl.Substring(pathStart).ToLowerInvariant();
                        }
                        return string.Empty;
                    }
                }
            }

            // Last resort: return the full path
            return lowerUrl;
        }
        catch
        {
            return serverRelativeUrl.ToLowerInvariant();
        }
    }

    /// <summary>
    /// Extracts the server-relative path from a full site URL.
    /// e.g., "https://tenant.sharepoint.com/sites/MySite" → "/sites/MySite"
    /// e.g., "https://tenant.sharepoint.com/" → "/"
    /// </summary>
    private static string GetServerRelativePath(string siteUrl)
    {
        try
        {
            var uri = new Uri(siteUrl.TrimEnd('/'));
            return uri.AbsolutePath;
        }
        catch
        {
            return string.Empty;
        }
    }

    #endregion

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

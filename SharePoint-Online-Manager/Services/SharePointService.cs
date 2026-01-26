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

    #region Permission Methods

    public async Task<SharePointResult<List<PermissionReportItem>>> GetWebPermissionsAsync(
        string siteUrl,
        string siteCollectionUrl,
        bool includeInherited = false)
    {
        try
        {
            var permissions = new List<PermissionReportItem>();
            var baseUrl = siteUrl.TrimEnd('/');

            // Check if web has unique permissions
            var hasUniqueUrl = $"{baseUrl}/_api/web?$select=Title,Url,HasUniqueRoleAssignments,ServerRelativeUrl";
            var webResponse = await _client.GetAsync(hasUniqueUrl);

            if (!webResponse.IsSuccessStatusCode)
            {
                return new SharePointResult<List<PermissionReportItem>>
                {
                    Status = GetSharePointStatus(webResponse.StatusCode),
                    ErrorMessage = GetErrorMessage(webResponse.StatusCode)
                };
            }

            var webJson = await webResponse.Content.ReadAsStringAsync();
            using var webDoc = JsonDocument.Parse(webJson);
            var webRoot = webDoc.RootElement;

            var webTitle = GetStringProperty(webRoot, "Title");
            var webUrl = GetStringProperty(webRoot, "Url");
            var serverRelativeUrl = GetStringProperty(webRoot, "ServerRelativeUrl");
            var hasUniquePerms = GetBoolProperty(webRoot, "HasUniqueRoleAssignments");

            // Determine object type
            var objectType = siteUrl.Equals(siteCollectionUrl, StringComparison.OrdinalIgnoreCase)
                ? PermissionObjectType.SiteCollection
                : (serverRelativeUrl.Count(c => c == '/') > 2 ? PermissionObjectType.Subsite : PermissionObjectType.Site);

            // If inherited and we don't want inherited, skip
            if (!hasUniquePerms && !includeInherited)
            {
                return new SharePointResult<List<PermissionReportItem>>
                {
                    Data = permissions,
                    Status = SharePointResultStatus.Success
                };
            }

            // Get role assignments
            var roleUrl = $"{baseUrl}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings";
            var roleResponse = await _client.GetAsync(roleUrl);

            if (roleResponse.IsSuccessStatusCode)
            {
                var roleJson = await roleResponse.Content.ReadAsStringAsync();
                var roleItems = ParseRoleAssignments(roleJson, siteCollectionUrl, siteUrl, webTitle,
                    objectType, webTitle, webUrl, !hasUniquePerms, siteUrl);
                permissions.AddRange(roleItems);
            }

            return new SharePointResult<List<PermissionReportItem>>
            {
                Data = permissions,
                Status = SharePointResultStatus.Success
            };
        }
        catch (Exception ex)
        {
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
        try
        {
            var permissions = new List<PermissionReportItem>();
            var baseUrl = siteUrl.TrimEnd('/');
            var encodedListTitle = Uri.EscapeDataString(listTitle);

            // Get list info including HasUniqueRoleAssignments
            var listUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedListTitle}')?$select=Title,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl&$expand=RootFolder";
            var listResponse = await _client.GetAsync(listUrl);

            if (!listResponse.IsSuccessStatusCode)
            {
                return new SharePointResult<List<PermissionReportItem>>
                {
                    Status = GetSharePointStatus(listResponse.StatusCode),
                    ErrorMessage = GetErrorMessage(listResponse.StatusCode)
                };
            }

            var listJson = await listResponse.Content.ReadAsStringAsync();
            using var listDoc = JsonDocument.Parse(listJson);
            var listRoot = listDoc.RootElement;

            var hasUniquePerms = GetBoolProperty(listRoot, "HasUniqueRoleAssignments");
            var listServerRelativeUrl = string.Empty;
            if (listRoot.TryGetProperty("RootFolder", out var rootFolder))
            {
                listServerRelativeUrl = GetStringProperty(rootFolder, "ServerRelativeUrl");
            }

            var siteUri = new Uri(siteUrl);
            var listAbsoluteUrl = !string.IsNullOrEmpty(listServerRelativeUrl)
                ? $"{siteUri.Scheme}://{siteUri.Host}{listServerRelativeUrl}"
                : siteUrl;

            // If inherited and we don't want inherited, skip
            if (!hasUniquePerms && !includeInherited)
            {
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
            var roleResponse = await _client.GetAsync(roleUrl);

            if (roleResponse.IsSuccessStatusCode)
            {
                var roleJson = await roleResponse.Content.ReadAsStringAsync();
                var roleItems = ParseRoleAssignments(roleJson, siteCollectionUrl, siteUrl, siteTitle,
                    objectType, listTitle, listAbsoluteUrl, !hasUniquePerms, siteUrl);
                permissions.AddRange(roleItems);
            }

            return new SharePointResult<List<PermissionReportItem>>
            {
                Data = permissions,
                Status = SharePointResultStatus.Success
            };
        }
        catch (Exception ex)
        {
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
        bool includeInherited = false)
    {
        try
        {
            var permissions = new List<PermissionReportItem>();
            var baseUrl = siteUrl.TrimEnd('/');
            var encodedListTitle = Uri.EscapeDataString(listTitle);

            // Get site title
            var siteInfo = await GetSiteInfoAsync(siteUrl);
            var siteTitle = siteInfo.Title;

            var siteUri = new Uri(siteUrl);

            // Fetch all items with HasUniqueRoleAssignments in $select, then filter client-side
            // This approach is more reliable than server-side $filter which isn't always supported
            var selectFields = "Id,HasUniqueRoleAssignments,FileRef,FileLeafRef,Title,FileSystemObjectType";
            int lastId = 0;
            bool hasMore = true;

            while (hasMore)
            {
                var apiUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedListTitle}')/items" +
                            $"?$select={selectFields}&$top=5000&$orderby=Id asc&$filter=Id gt {lastId}";

                var response = await _client.GetAsync(apiUrl);

                if (!response.IsSuccessStatusCode)
                {
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
                    break;
                }

                int itemCount = 0;
                foreach (var item in valueElement.EnumerateArray())
                {
                    itemCount++;
                    var itemId = GetIntProperty(item, "Id");
                    lastId = itemId;

                    // Check if item has unique permissions - filter client-side
                    var hasUniquePerms = GetBoolProperty(item, "HasUniqueRoleAssignments");
                    if (!hasUniquePerms)
                        continue;

                    var fileRef = GetStringProperty(item, "FileRef");
                    var fileName = GetStringProperty(item, "FileLeafRef");
                    // FileSystemObjectType: 0 = file, 1 = folder
                    var fsObjType = GetIntProperty(item, "FileSystemObjectType");

                    // Use Title as fallback for list items without FileLeafRef
                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = GetStringProperty(item, "Title");
                    }

                    var isFolder = fsObjType == 1;

                    // Skip based on options
                    if (isFolder && !includeFolders) continue;
                    if (!isFolder && !includeItems) continue;

                    var absoluteUrl = !string.IsNullOrEmpty(fileRef)
                        ? $"{siteUri.Scheme}://{siteUri.Host}{fileRef}"
                        : siteUrl;
                    var objectType = isFolder ? PermissionObjectType.Folder
                        : (isLibrary ? PermissionObjectType.Document : PermissionObjectType.ListItem);

                    // Get role assignments for this item
                    var roleUrl = $"{baseUrl}/_api/web/lists/GetByTitle('{encodedListTitle}')/items({itemId})/roleassignments?$expand=Member,RoleDefinitionBindings";
                    var roleResponse = await _client.GetAsync(roleUrl);

                    if (roleResponse.IsSuccessStatusCode)
                    {
                        var roleJson = await roleResponse.Content.ReadAsStringAsync();
                        var roleItems = ParseRoleAssignments(roleJson, siteCollectionUrl, siteUrl, siteTitle,
                            objectType, fileName, absoluteUrl, false, siteUrl);

                        // Set the object path
                        foreach (var roleItem in roleItems)
                        {
                            roleItem.ObjectPath = fileRef;
                        }

                        permissions.AddRange(roleItems);
                    }
                }

                // If we got fewer than 5000 items, we've reached the end
                hasMore = itemCount == 5000;
            }

            return new SharePointResult<List<PermissionReportItem>>
            {
                Data = permissions,
                Status = SharePointResultStatus.Success
            };
        }
        catch (Exception ex)
        {
            return new SharePointResult<List<PermissionReportItem>>
            {
                Status = SharePointResultStatus.Error,
                ErrorMessage = ex.Message
            };
        }
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
                return items;
            }

            foreach (var assignment in valueElement.EnumerateArray())
            {
                // Get member info
                if (!assignment.TryGetProperty("Member", out var member))
                    continue;

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
                    continue;

                JsonElement rolesArray;
                if (roleBindings.TryGetProperty("results", out rolesArray))
                {
                    // OData verbose
                }
                else if (roleBindings.ValueKind == JsonValueKind.Array)
                {
                    rolesArray = roleBindings;
                }
                else
                {
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
                    continue;

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
        }
        catch { }

        return items;
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

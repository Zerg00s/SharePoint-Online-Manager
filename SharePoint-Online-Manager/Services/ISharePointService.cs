using SharePointOnlineManager.Models;
using SharePointOnlineManager.Screens;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Indicates the result status of a SharePoint API call.
/// </summary>
public enum SharePointResultStatus
{
    Success,
    AuthenticationRequired,
    AccessDenied,
    NotFound,
    Error
}

/// <summary>
/// Result wrapper for SharePoint API calls with authentication status.
/// </summary>
/// <typeparam name="T">The result data type.</typeparam>
public class SharePointResult<T>
{
    public T? Data { get; init; }
    public SharePointResultStatus Status { get; init; }
    public string? ErrorMessage { get; init; }
    public bool IsSuccess => Status == SharePointResultStatus.Success;
    public bool NeedsReauth => Status == SharePointResultStatus.AuthenticationRequired;
}

/// <summary>
/// Interface for SharePoint REST API operations.
/// </summary>
public interface ISharePointService : IDisposable
{
    /// <summary>
    /// Gets the domain this service is authenticated for.
    /// </summary>
    string Domain { get; }

    /// <summary>
    /// Gets information about a SharePoint site.
    /// </summary>
    Task<SiteInfo> GetSiteInfoAsync(string siteUrl);

    /// <summary>
    /// Tests the connection to a SharePoint site.
    /// </summary>
    Task<bool> TestConnectionAsync(string siteUrl);

    /// <summary>
    /// Gets all lists from a SharePoint site.
    /// </summary>
    Task<SharePointResult<List<ListInfo>>> GetListsAsync(string siteUrl);

    /// <summary>
    /// Gets all lists from a SharePoint site with optional filtering.
    /// </summary>
    Task<SharePointResult<List<ListInfo>>> GetListsAsync(string siteUrl, bool includeHidden);

    /// <summary>
    /// Gets all subsites (subwebs) of a SharePoint site.
    /// </summary>
    Task<List<SubsiteInfo>> GetSubsitesAsync(string siteUrl);

    /// <summary>
    /// Gets all files from a document library.
    /// </summary>
    /// <param name="siteUrl">The SharePoint site URL.</param>
    /// <param name="libraryTitle">The document library title.</param>
    /// <param name="includeSubfolders">Whether to include files from subfolders.</param>
    /// <param name="includeVersionCount">Whether to retrieve version count for each file.</param>
    /// <returns>A list of document report items.</returns>
    Task<SharePointResult<List<DocumentReportItem>>> GetDocumentLibraryFilesAsync(
        string siteUrl,
        string libraryTitle,
        bool includeSubfolders = true,
        bool includeVersionCount = true);

    /// <summary>
    /// Gets permission role assignments for a site/web.
    /// </summary>
    Task<SharePointResult<List<PermissionReportItem>>> GetWebPermissionsAsync(
        string siteUrl,
        string siteCollectionUrl,
        bool includeInherited = false);

    /// <summary>
    /// Gets permission role assignments for a list or library.
    /// </summary>
    Task<SharePointResult<List<PermissionReportItem>>> GetListPermissionsAsync(
        string siteUrl,
        string siteCollectionUrl,
        string listTitle,
        bool isLibrary,
        bool includeInherited = false);

    /// <summary>
    /// Gets items/folders with unique permissions in a list or library.
    /// </summary>
    Task<SharePointResult<List<PermissionReportItem>>> GetItemPermissionsAsync(
        string siteUrl,
        string siteCollectionUrl,
        string listTitle,
        bool isLibrary,
        bool includeFolders = true,
        bool includeItems = true,
        bool includeInherited = false);

    /// <summary>
    /// Deletes a site collection (sends to recycle bin).
    /// Requires SharePoint Admin or Site Collection Administrator permissions.
    /// </summary>
    /// <param name="siteUrl">The URL of the site collection to delete.</param>
    /// <returns>A result indicating success or failure.</returns>
    Task<SharePointResult<bool>> DeleteSiteCollectionAsync(string siteUrl);
}

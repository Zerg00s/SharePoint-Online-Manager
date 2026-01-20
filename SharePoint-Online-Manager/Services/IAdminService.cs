using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Interface for SharePoint Online Admin API operations.
/// </summary>
public interface IAdminService : IDisposable
{
    /// <summary>
    /// Gets all site collections from the tenant.
    /// </summary>
    /// <param name="progress">Optional progress reporter for UI updates.</param>
    Task<List<SiteCollection>> GetAllSiteCollectionsAsync(IProgress<string>? progress = null);

    /// <summary>
    /// Gets site collections filtered by type.
    /// </summary>
    Task<List<SiteCollection>> GetSiteCollectionsByTypeAsync(SiteType type, IProgress<string>? progress = null);

    /// <summary>
    /// Tests the admin API connection.
    /// </summary>
    Task<bool> TestConnectionAsync();
}

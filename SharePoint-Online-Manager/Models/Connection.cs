namespace SharePointOnlineManager.Models;

/// <summary>
/// Defines the type of SharePoint connection.
/// </summary>
public enum ConnectionType
{
    /// <summary>
    /// Admin connection for tenant-level operations (e.g., contoso-admin.sharepoint.com).
    /// </summary>
    Admin,

    /// <summary>
    /// Direct connection to a specific site collection.
    /// </summary>
    SiteCollection
}

/// <summary>
/// Represents a saved SharePoint connection configuration.
/// </summary>
public class Connection
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = string.Empty;
    public ConnectionType Type { get; set; }
    public string TenantName { get; set; } = string.Empty;
    public string? SiteUrl { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime? LastConnectedAt { get; set; }

    /// <summary>
    /// Gets the admin URL for this connection's tenant.
    /// </summary>
    public string AdminUrl => $"https://{TenantName}-admin.sharepoint.com";

    /// <summary>
    /// Gets the regular tenant URL (e.g., contoso.sharepoint.com).
    /// </summary>
    public string TenantUrl => $"https://{TenantName}.sharepoint.com";

    /// <summary>
    /// Gets the regular tenant domain (e.g., contoso.sharepoint.com).
    /// </summary>
    public string TenantDomain => $"{TenantName}.sharepoint.com";

    /// <summary>
    /// Gets the admin domain (e.g., contoso-admin.sharepoint.com).
    /// </summary>
    public string AdminDomain => $"{TenantName}-admin.sharepoint.com";

    /// <summary>
    /// Gets the domain for cookie storage based on connection type.
    /// For Admin connections, uses the admin domain to access the aggregated sites list.
    /// </summary>
    public string CookieDomain => Type == ConnectionType.Admin
        ? AdminDomain
        : new Uri(SiteUrl ?? TenantUrl).Host;

    /// <summary>
    /// Gets the primary URL for this connection (used for authentication).
    /// For Admin connections, uses the admin URL to access tenant-level data.
    /// </summary>
    public string PrimaryUrl => Type == ConnectionType.Admin
        ? AdminUrl
        : SiteUrl ?? TenantUrl;
}

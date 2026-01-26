namespace SharePointOnlineManager.Models;

/// <summary>
/// Defines the type of SharePoint object for permission reporting.
/// </summary>
public enum PermissionObjectType
{
    SiteCollection,
    Site,
    Subsite,
    List,
    Library,
    Folder,
    ListItem,
    Document
}

/// <summary>
/// Configuration for a permission report task.
/// </summary>
public class PermissionReportConfiguration
{
    public Guid ConnectionId { get; set; }
    public List<string> TargetSiteUrls { get; set; } = [];
    public bool IncludeSitePermissions { get; set; } = true;
    public bool IncludeListPermissions { get; set; } = true;
    public bool IncludeFolderPermissions { get; set; } = true;
    public bool IncludeItemPermissions { get; set; } = false; // Can be very slow for large libraries
    public bool IncludeInheritedPermissions { get; set; } = false; // Only show unique permissions by default
    public bool IncludeHiddenLists { get; set; } = false;
}

/// <summary>
/// Represents a single permission entry in the report.
/// </summary>
public class PermissionReportItem
{
    public string SiteCollectionUrl { get; set; } = string.Empty;
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public PermissionObjectType ObjectType { get; set; }
    public string ObjectTypeDescription => ObjectType switch
    {
        PermissionObjectType.SiteCollection => "Site Collection",
        PermissionObjectType.Site => "Site",
        PermissionObjectType.Subsite => "Subsite",
        PermissionObjectType.List => "List",
        PermissionObjectType.Library => "Library",
        PermissionObjectType.Folder => "Folder",
        PermissionObjectType.ListItem => "List Item",
        PermissionObjectType.Document => "Document",
        _ => ObjectType.ToString()
    };
    public string ObjectTitle { get; set; } = string.Empty;
    public string ObjectUrl { get; set; } = string.Empty;
    public string ObjectPath { get; set; } = string.Empty;
    public string PrincipalName { get; set; } = string.Empty;
    public string PrincipalType { get; set; } = string.Empty; // User, Group, SharePointGroup
    public string PrincipalLogin { get; set; } = string.Empty;
    public string PermissionLevel { get; set; } = string.Empty;
    public bool IsInherited { get; set; }
    public string InheritedFrom { get; set; } = string.Empty;
}

/// <summary>
/// Represents permission results for a single SharePoint site.
/// </summary>
public class SitePermissionResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<PermissionReportItem> Permissions { get; set; } = [];
    public int TotalPermissions => Permissions.Count;
    public int UniquePermissionObjects => Permissions.Select(p => p.ObjectUrl).Distinct().Count();
}

/// <summary>
/// Represents the complete result of a permission report task execution.
/// </summary>
public class PermissionReportResult
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public Guid TaskId { get; set; }
    public DateTime ExecutedAt { get; set; } = DateTime.UtcNow;
    public DateTime? CompletedAt { get; set; }
    public TimeSpan Duration => (CompletedAt ?? DateTime.UtcNow) - ExecutedAt;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public int TotalSitesProcessed { get; set; }
    public int SuccessfulSites { get; set; }
    public int FailedSites { get; set; }
    public List<SitePermissionResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets all permissions flattened across all sites.
    /// </summary>
    public IEnumerable<PermissionReportItem> GetAllPermissions()
    {
        foreach (var siteResult in SiteResults)
        {
            foreach (var perm in siteResult.Permissions)
            {
                yield return perm;
            }
        }
    }

    /// <summary>
    /// Gets summary statistics for the report.
    /// </summary>
    public (int totalPermissions, int uniqueObjects, int uniquePrincipals) GetSummary()
    {
        var allPerms = GetAllPermissions().ToList();
        var totalPerms = allPerms.Count;
        var uniqueObjects = allPerms.Select(p => p.ObjectUrl).Distinct().Count();
        var uniquePrincipals = allPerms.Select(p => p.PrincipalName).Distinct().Count();
        return (totalPerms, uniqueObjects, uniquePrincipals);
    }

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }
}

/// <summary>
/// Export model for permission report CSV.
/// </summary>
public class PermissionReportExportItem
{
    public string ObjectType { get; set; } = string.Empty;
    public string SiteCollectionUrl { get; set; } = string.Empty;
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string ObjectTitle { get; set; } = string.Empty;
    public string ObjectUrl { get; set; } = string.Empty;
    public string ObjectPath { get; set; } = string.Empty;
    public string PrincipalName { get; set; } = string.Empty;
    public string PrincipalType { get; set; } = string.Empty;
    public string PrincipalLogin { get; set; } = string.Empty;
    public string PermissionLevel { get; set; } = string.Empty;
    public string IsInherited { get; set; } = string.Empty;
    public string InheritedFrom { get; set; } = string.Empty;
}

/// <summary>
/// Export model for permission report site summary.
/// </summary>
public class PermissionReportSiteSummaryExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public int TotalPermissions { get; set; }
    public int UniqueObjects { get; set; }
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

namespace SharePointOnlineManager.Models;

/// <summary>
/// Configuration for an ad hoc users report task.
/// </summary>
public class AdHocUsersReportConfiguration
{
    public Guid ConnectionId { get; set; }
    public List<string> TargetSiteUrls { get; set; } = [];
}

/// <summary>
/// Represents a single ad hoc (OTP) guest user found on a site.
/// </summary>
public class AdHocUserItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string LoginName { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public int Id { get; set; }
    public bool IsSiteAdmin { get; set; }
    public string PrincipalType { get; set; } = string.Empty;
}

/// <summary>
/// Represents ad hoc user results for a single SharePoint site.
/// </summary>
public class SiteAdHocUsersResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<AdHocUserItem> Users { get; set; } = [];
    public int GuestCount => Users.Count;
}

/// <summary>
/// Represents the complete result of an ad hoc users report task execution.
/// </summary>
public class AdHocUsersReportResult
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
    public List<SiteAdHocUsersResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets all ad hoc users flattened across all sites.
    /// </summary>
    public IEnumerable<AdHocUserItem> GetAllUsers()
    {
        foreach (var siteResult in SiteResults)
        {
            foreach (var user in siteResult.Users)
            {
                yield return user;
            }
        }
    }

    /// <summary>
    /// Gets the total number of guest users across all sites.
    /// </summary>
    public int TotalGuestUsers => SiteResults.Sum(s => s.GuestCount);

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }
}

/// <summary>
/// Export model for ad hoc users report CSV.
/// </summary>
public class AdHocUsersExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string LoginName { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public int Id { get; set; }
    public string IsSiteAdmin { get; set; } = string.Empty;
    public string PrincipalType { get; set; } = string.Empty;
}

/// <summary>
/// Export model for ad hoc users site summary CSV.
/// </summary>
public class AdHocUsersSiteSummaryExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public int GuestCount { get; set; }
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

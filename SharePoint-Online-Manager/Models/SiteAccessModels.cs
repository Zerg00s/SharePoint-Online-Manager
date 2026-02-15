namespace SharePointOnlineManager.Models;

/// <summary>
/// Access status for a site.
/// </summary>
public enum SiteAccessStatus
{
    /// <summary>
    /// User has access to the site.
    /// </summary>
    Accessible,

    /// <summary>
    /// Access denied (403 Forbidden).
    /// </summary>
    AccessDenied,

    /// <summary>
    /// Site not found (404).
    /// </summary>
    NotFound,

    /// <summary>
    /// Authentication required (401).
    /// </summary>
    AuthenticationRequired,

    /// <summary>
    /// Other error occurred.
    /// </summary>
    Error
}

/// <summary>
/// Configuration for a site access check task.
/// </summary>
public class SiteAccessConfiguration
{
    public Guid SourceConnectionId { get; set; }
    public Guid TargetConnectionId { get; set; }
    public List<SiteComparePair> SitePairs { get; set; } = [];
}

/// <summary>
/// Result of checking access to a single site.
/// </summary>
public class SiteAccessCheckItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public SiteAccessStatus Status { get; set; }
    public string? ErrorMessage { get; set; }
    public string? AccountUsed { get; set; }
    public bool IsSource { get; set; }

    /// <summary>
    /// Gets a display string for the status.
    /// </summary>
    public string StatusDescription => Status switch
    {
        SiteAccessStatus.Accessible => "Accessible",
        SiteAccessStatus.AccessDenied => "Access Denied",
        SiteAccessStatus.NotFound => "Not Found",
        SiteAccessStatus.AuthenticationRequired => "Auth Required",
        SiteAccessStatus.Error => "Error",
        _ => Status.ToString()
    };

    /// <summary>
    /// Indicates if this is an access issue.
    /// </summary>
    public bool HasIssue => Status != SiteAccessStatus.Accessible;
}

/// <summary>
/// Result of checking access to a site pair (source and target).
/// </summary>
public class SitePairAccessResult
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public SiteAccessCheckItem SourceResult { get; set; } = new();
    public SiteAccessCheckItem TargetResult { get; set; } = new();

    /// <summary>
    /// Indicates if either source or target has an access issue.
    /// </summary>
    public bool HasIssue => SourceResult.HasIssue || TargetResult.HasIssue;

    /// <summary>
    /// Indicates if source has an access issue.
    /// </summary>
    public bool HasSourceIssue => SourceResult.HasIssue;

    /// <summary>
    /// Indicates if target has an access issue.
    /// </summary>
    public bool HasTargetIssue => TargetResult.HasIssue;
}

/// <summary>
/// Complete result of a site access check task.
/// </summary>
public class SiteAccessResult
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public Guid TaskId { get; set; }
    public DateTime ExecutedAt { get; set; } = DateTime.UtcNow;
    public DateTime? CompletedAt { get; set; }
    public TimeSpan Duration => (CompletedAt ?? DateTime.UtcNow) - ExecutedAt;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }

    public string? SourceAccount { get; set; }
    public string? TargetAccount { get; set; }

    public int TotalPairsProcessed { get; set; }
    public List<SitePairAccessResult> PairResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets count of source sites with access.
    /// </summary>
    public int SourceAccessibleCount => PairResults.Count(p => !p.SourceResult.HasIssue);

    /// <summary>
    /// Gets count of source sites with access denied.
    /// </summary>
    public int SourceAccessDeniedCount => PairResults.Count(p => p.SourceResult.Status == SiteAccessStatus.AccessDenied);

    /// <summary>
    /// Gets count of source sites with other issues.
    /// </summary>
    public int SourceOtherIssuesCount => PairResults.Count(p =>
        p.SourceResult.HasIssue && p.SourceResult.Status != SiteAccessStatus.AccessDenied);

    /// <summary>
    /// Gets count of target sites with access.
    /// </summary>
    public int TargetAccessibleCount => PairResults.Count(p => !p.TargetResult.HasIssue);

    /// <summary>
    /// Gets count of target sites with access denied.
    /// </summary>
    public int TargetAccessDeniedCount => PairResults.Count(p => p.TargetResult.Status == SiteAccessStatus.AccessDenied);

    /// <summary>
    /// Gets count of target sites with other issues.
    /// </summary>
    public int TargetOtherIssuesCount => PairResults.Count(p =>
        p.TargetResult.HasIssue && p.TargetResult.Status != SiteAccessStatus.AccessDenied);

    /// <summary>
    /// Gets all source sites with issues.
    /// </summary>
    public IEnumerable<SiteAccessCheckItem> GetSourceIssues() =>
        PairResults.Where(p => p.HasSourceIssue).Select(p => p.SourceResult);

    /// <summary>
    /// Gets all target sites with issues.
    /// </summary>
    public IEnumerable<SiteAccessCheckItem> GetTargetIssues() =>
        PairResults.Where(p => p.HasTargetIssue).Select(p => p.TargetResult);

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }
}

/// <summary>
/// Export model for site access check results (CSV).
/// </summary>
public class SiteAccessExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string Side { get; set; } = string.Empty; // "Source" or "Target"
    public string Status { get; set; } = string.Empty;
    public string AccountUsed { get; set; } = string.Empty;
    public string ErrorMessage { get; set; } = string.Empty;
}

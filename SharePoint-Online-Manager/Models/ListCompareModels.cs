namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents a source/target site URL pair for list comparison.
/// </summary>
public class SiteComparePair
{
    public string SourceUrl { get; set; } = string.Empty;
    public string TargetUrl { get; set; } = string.Empty;
}

/// <summary>
/// Defines how threshold comparisons are measured.
/// </summary>
public enum ThresholdType
{
    /// <summary>
    /// Threshold is a percentage difference.
    /// </summary>
    Percentage,

    /// <summary>
    /// Threshold is an absolute item count difference.
    /// </summary>
    AbsoluteCount
}

/// <summary>
/// Configuration for a list compare task.
/// </summary>
public class ListCompareConfiguration
{
    public Guid SourceConnectionId { get; set; }
    public Guid TargetConnectionId { get; set; }
    public List<SiteComparePair> SitePairs { get; set; } = [];
    public List<string> ExcludedLists { get; set; } = [];
    public bool IncludeSiteAssets { get; set; } = false;
    public bool IncludeHiddenLists { get; set; } = false;
    public ThresholdType ThresholdType { get; set; } = ThresholdType.Percentage;
    public int ThresholdValue { get; set; } = 10;

    /// <summary>
    /// Gets the default list of system lists that should always be excluded.
    /// </summary>
    public static List<string> DefaultExcludedLists =>
    [
        "MicroFeed",
        "Style Library",
        "appdata",
        "TaxonomyHiddenList",
        "Composed Looks",
        "Master Page Gallery",
        "Solution Gallery",
        "Theme Gallery",
        "Web Part Gallery",
        "Workflow Tasks",
        "User Information List",
        "Converted Forms",
        "Customized Reports",
        "Form Templates",
        "Content type publishing error log",
        "Team Message History",
        "Channel Settings"
    ];

    /// <summary>
    /// Gets lists that are excluded by default but user can choose to include.
    /// </summary>
    public static List<string> OptionalExcludedLists =>
    [
        "Site Assets"
    ];

    /// <summary>
    /// Gets lists that should never be excluded.
    /// </summary>
    public static List<string> NeverExcludedLists =>
    [
        "Site Pages"
    ];
}

/// <summary>
/// Defines the comparison status for a list.
/// </summary>
public enum ListCompareStatus
{
    /// <summary>
    /// Item counts match (within threshold).
    /// </summary>
    Match,

    /// <summary>
    /// Item counts differ (beyond threshold).
    /// </summary>
    Mismatch,

    /// <summary>
    /// List exists only on source site.
    /// </summary>
    SourceOnly,

    /// <summary>
    /// List exists only on target site.
    /// </summary>
    TargetOnly
}

/// <summary>
/// Represents a single list comparison result.
/// </summary>
public class ListCompareItem
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string SourceSiteTitle { get; set; } = string.Empty;
    public string TargetSiteTitle { get; set; } = string.Empty;
    public string ListTitle { get; set; } = string.Empty;
    public string ListType { get; set; } = string.Empty;
    public string SourceListUrl { get; set; } = string.Empty;
    public string TargetListUrl { get; set; } = string.Empty;
    public int SourceCount { get; set; }
    public int TargetCount { get; set; }
    public int Difference => TargetCount - SourceCount;
    public double PercentDifference => SourceCount == 0
        ? (TargetCount == 0 ? 0 : 100)
        : Math.Abs((double)(TargetCount - SourceCount) / SourceCount * 100);
    public ListCompareStatus Status { get; set; }

    /// <summary>
    /// Gets a display string for the status.
    /// </summary>
    public string StatusDescription => Status switch
    {
        ListCompareStatus.Match => "Match",
        ListCompareStatus.Mismatch => "Mismatch",
        ListCompareStatus.SourceOnly => "Source Only",
        ListCompareStatus.TargetOnly => "Target Only",
        _ => Status.ToString()
    };
}

/// <summary>
/// Represents comparison results for a single site pair.
/// </summary>
public class SiteCompareResult
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string SourceSiteTitle { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string TargetSiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<ListCompareItem> ListComparisons { get; set; } = [];

    /// <summary>
    /// Gets the count of lists that match.
    /// </summary>
    public int MatchCount => ListComparisons.Count(l => l.Status == ListCompareStatus.Match);

    /// <summary>
    /// Gets the count of lists that mismatch.
    /// </summary>
    public int MismatchCount => ListComparisons.Count(l => l.Status == ListCompareStatus.Mismatch);

    /// <summary>
    /// Gets the count of lists that exist only on source.
    /// </summary>
    public int SourceOnlyCount => ListComparisons.Count(l => l.Status == ListCompareStatus.SourceOnly);

    /// <summary>
    /// Gets the count of lists that exist only on target.
    /// </summary>
    public int TargetOnlyCount => ListComparisons.Count(l => l.Status == ListCompareStatus.TargetOnly);

    /// <summary>
    /// Indicates whether this site pair has any issues.
    /// </summary>
    public bool HasIssues => !Success || MismatchCount > 0 || SourceOnlyCount > 0 || TargetOnlyCount > 0;
}

/// <summary>
/// Represents the complete result of a list compare task execution.
/// </summary>
public class ListCompareResult
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public Guid TaskId { get; set; }
    public DateTime ExecutedAt { get; set; } = DateTime.UtcNow;
    public DateTime? CompletedAt { get; set; }
    public TimeSpan Duration => (CompletedAt ?? DateTime.UtcNow) - ExecutedAt;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public int TotalPairsProcessed { get; set; }
    public int SuccessfulPairs { get; set; }
    public int FailedPairs { get; set; }
    public List<SiteCompareResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets all list comparison items flattened across all site pairs.
    /// </summary>
    public IEnumerable<ListCompareItem> GetAllListComparisons()
    {
        foreach (var siteResult in SiteResults)
        {
            foreach (var comparison in siteResult.ListComparisons)
            {
                yield return comparison;
            }
        }
    }

    /// <summary>
    /// Gets site results that have issues (errors, mismatches, or missing lists).
    /// </summary>
    public IEnumerable<SiteCompareResult> GetSitesWithIssues()
    {
        return SiteResults.Where(s => s.HasIssues);
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
/// Export model for list comparison results.
/// </summary>
public class ListCompareExportItem
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string ListTitle { get; set; } = string.Empty;
    public string ListType { get; set; } = string.Empty;
    public int SourceCount { get; set; }
    public int TargetCount { get; set; }
    public int Difference { get; set; }
    public string PercentDifference { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}

/// <summary>
/// Export model for site issues summary.
/// </summary>
public class SiteIssueExportItem
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public int Mismatches { get; set; }
    public int SourceOnly { get; set; }
    public int TargetOnly { get; set; }
    public string Error { get; set; } = string.Empty;
}

/// <summary>
/// Export model for list mapping with full URLs.
/// </summary>
public class ListMappingExportItem
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string SourceSiteTitle { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string TargetSiteTitle { get; set; } = string.Empty;
    public string ListTitle { get; set; } = string.Empty;
    public string ListType { get; set; } = string.Empty;
    public string SourceListUrl { get; set; } = string.Empty;
    public string TargetListUrl { get; set; } = string.Empty;
    public int SourceCount { get; set; }
    public int TargetCount { get; set; }
    public string Status { get; set; } = string.Empty;
}

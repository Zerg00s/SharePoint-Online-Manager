namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents the results from executing a task.
/// </summary>
public class TaskResult
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
    public List<SiteListResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets all list items flattened across all sites.
    /// </summary>
    public IEnumerable<ListReportItem> GetAllListItems()
    {
        foreach (var siteResult in SiteResults)
        {
            foreach (var list in siteResult.Lists)
            {
                yield return new ListReportItem
                {
                    SiteUrl = siteResult.SiteUrl,
                    SiteTitle = siteResult.SiteTitle,
                    ListId = list.Id,
                    ListTitle = list.Title,
                    ListUrl = list.GetAbsoluteUrl(siteResult.SiteUrl),
                    ItemCount = list.ItemCount,
                    Hidden = list.Hidden,
                    ListType = list.ListType,
                    Created = list.Created,
                    LastModified = list.LastItemModifiedDate
                };
            }
        }
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
/// Represents the results for a single site within a task execution.
/// </summary>
public class SiteListResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<ListInfo> Lists { get; set; } = [];
}

/// <summary>
/// Represents a flattened list item for reporting purposes.
/// </summary>
public class ListReportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public Guid ListId { get; set; }
    public string ListTitle { get; set; } = string.Empty;
    public string ListUrl { get; set; } = string.Empty;
    public int ItemCount { get; set; }
    public bool Hidden { get; set; }
    public string ListType { get; set; } = string.Empty;
    public DateTime Created { get; set; }
    public DateTime LastModified { get; set; }
}

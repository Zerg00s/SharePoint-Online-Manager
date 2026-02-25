namespace SharePointOnlineManager.Models;

/// <summary>
/// Well-known SharePoint Publishing feature GUIDs.
/// </summary>
public static class PublishingFeatureIds
{
    /// <summary>
    /// SharePoint Server Publishing Infrastructure (Site Collection feature).
    /// </summary>
    public static readonly Guid PublishingInfrastructure = new("f6924d36-2fa8-4f0b-b16d-06b7250180fa");

    /// <summary>
    /// SharePoint Server Publishing (Web/Site feature).
    /// </summary>
    public static readonly Guid PublishingWeb = new("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb");
}

/// <summary>
/// Represents the publishing feature status for a single SharePoint site.
/// </summary>
public class SitePublishingResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public bool HasPublishingInfrastructure { get; set; }
    public bool HasPublishingWeb { get; set; }
    public bool HasPublishing => HasPublishingInfrastructure || HasPublishingWeb;

    public string PublishingStatus => (HasPublishingInfrastructure, HasPublishingWeb) switch
    {
        (true, true) => "Both Active",
        (true, false) => "Infrastructure Only",
        (false, true) => "Web Only",
        (false, false) => "Not Active"
    };
}

/// <summary>
/// Represents the complete result of a publishing sites report task execution.
/// </summary>
public class PublishingSitesReportResult
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
    public List<SitePublishingResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    public int PublishingSitesCount => SiteResults.Count(s => s.Success && s.HasPublishing);
    public int NonPublishingSitesCount => SiteResults.Count(s => s.Success && !s.HasPublishing);
    public int BothActiveCount => SiteResults.Count(s => s.Success && s.HasPublishingInfrastructure && s.HasPublishingWeb);
    public int InfraOnlyCount => SiteResults.Count(s => s.Success && s.HasPublishingInfrastructure && !s.HasPublishingWeb);
    public int WebOnlyCount => SiteResults.Count(s => s.Success && !s.HasPublishingInfrastructure && s.HasPublishingWeb);

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }
}

/// <summary>
/// Export model for publishing sites report CSV.
/// </summary>
public class PublishingSitesExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string PublishingInfrastructure { get; set; } = string.Empty;
    public string PublishingWeb { get; set; } = string.Empty;
    public string PublishingStatus { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

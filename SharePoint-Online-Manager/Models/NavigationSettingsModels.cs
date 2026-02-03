namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents navigation settings for a SharePoint site.
/// </summary>
public class NavigationSettings
{
    public bool HorizontalQuickLaunch { get; set; }
    public bool MegaMenuEnabled { get; set; }
}

/// <summary>
/// Configuration for a navigation settings sync task.
/// </summary>
public class NavigationSettingsConfiguration
{
    public Guid SourceConnectionId { get; set; }
    public Guid TargetConnectionId { get; set; }
    public List<SiteComparePair> SitePairs { get; set; } = [];
}

/// <summary>
/// Defines the comparison status for navigation settings.
/// </summary>
public enum NavigationSettingsStatus
{
    /// <summary>
    /// Settings match between source and target.
    /// </summary>
    Match,

    /// <summary>
    /// Settings differ between source and target.
    /// </summary>
    Mismatch,

    /// <summary>
    /// Settings were successfully applied to target.
    /// </summary>
    Applied,

    /// <summary>
    /// Failed to apply settings to target.
    /// </summary>
    Failed,

    /// <summary>
    /// Error accessing the site.
    /// </summary>
    Error
}

/// <summary>
/// Represents a single site's navigation settings comparison result.
/// </summary>
public class NavigationSettingsCompareItem
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string SourceSiteTitle { get; set; } = string.Empty;
    public string TargetSiteTitle { get; set; } = string.Empty;

    // Source settings
    public bool SourceHorizontalQuickLaunch { get; set; }
    public bool SourceMegaMenuEnabled { get; set; }

    // Target settings (before apply)
    public bool TargetHorizontalQuickLaunch { get; set; }
    public bool TargetMegaMenuEnabled { get; set; }

    public NavigationSettingsStatus Status { get; set; }
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// Indicates if HorizontalQuickLaunch matches between source and target.
    /// </summary>
    public bool HorizontalQuickLaunchMatches => SourceHorizontalQuickLaunch == TargetHorizontalQuickLaunch;

    /// <summary>
    /// Indicates if MegaMenuEnabled matches between source and target.
    /// </summary>
    public bool MegaMenuEnabledMatches => SourceMegaMenuEnabled == TargetMegaMenuEnabled;

    /// <summary>
    /// Indicates if all settings match.
    /// </summary>
    public bool AllSettingsMatch => HorizontalQuickLaunchMatches && MegaMenuEnabledMatches;

    /// <summary>
    /// Gets a display string for the status.
    /// </summary>
    public string StatusDescription => Status switch
    {
        NavigationSettingsStatus.Match => "Match",
        NavigationSettingsStatus.Mismatch => "Mismatch",
        NavigationSettingsStatus.Applied => "Applied",
        NavigationSettingsStatus.Failed => "Failed",
        NavigationSettingsStatus.Error => "Error",
        _ => Status.ToString()
    };
}

/// <summary>
/// Represents the complete result of a navigation settings sync task execution.
/// </summary>
public class NavigationSettingsResult
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public Guid TaskId { get; set; }
    public DateTime ExecutedAt { get; set; } = DateTime.UtcNow;
    public DateTime? CompletedAt { get; set; }
    public TimeSpan Duration => (CompletedAt ?? DateTime.UtcNow) - ExecutedAt;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public int TotalPairsProcessed { get; set; }
    public int MatchingPairs { get; set; }
    public int MismatchedPairs { get; set; }
    public int AppliedPairs { get; set; }
    public int FailedPairs { get; set; }
    public bool ApplyMode { get; set; }
    public List<NavigationSettingsCompareItem> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets site results that have mismatches.
    /// </summary>
    public IEnumerable<NavigationSettingsCompareItem> GetMismatchedSites()
    {
        return SiteResults.Where(s => s.Status == NavigationSettingsStatus.Mismatch);
    }

    /// <summary>
    /// Gets site results that had errors.
    /// </summary>
    public IEnumerable<NavigationSettingsCompareItem> GetFailedSites()
    {
        return SiteResults.Where(s => s.Status == NavigationSettingsStatus.Error || s.Status == NavigationSettingsStatus.Failed);
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
/// Export model for navigation settings comparison results.
/// </summary>
public class NavigationSettingsExportItem
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string SourceSiteTitle { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string TargetSiteTitle { get; set; } = string.Empty;
    public string SourceHorizontalNav { get; set; } = string.Empty;
    public string TargetHorizontalNav { get; set; } = string.Empty;
    public string HorizontalNavMatch { get; set; } = string.Empty;
    public string SourceMegaMenu { get; set; } = string.Empty;
    public string TargetMegaMenu { get; set; } = string.Empty;
    public string MegaMenuMatch { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

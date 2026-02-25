namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents a subsite found during the subsites report scan.
/// Extends SubsiteInfo with additional properties (language, size, etc.).
/// </summary>
public class SubsiteReportItem
{
    public string SiteCollectionUrl { get; set; } = string.Empty;
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string SubsiteUrl { get; set; } = string.Empty;
    public string SubsiteTitle { get; set; } = string.Empty;
    public string ServerRelativeUrl { get; set; } = string.Empty;
    public string WebTemplate { get; set; } = string.Empty;
    public DateTime Created { get; set; }
    public DateTime LastModified { get; set; }
    public int Language { get; set; }
    public string LanguageDisplay => Language switch
    {
        1033 => "English",
        1036 => "French",
        3084 => "French (Canada)",
        1031 => "German",
        1034 => "Spanish",
        1040 => "Italian",
        1041 => "Japanese",
        1042 => "Korean",
        1043 => "Dutch",
        1046 => "Portuguese (Brazil)",
        2070 => "Portuguese (Portugal)",
        1049 => "Russian",
        2052 => "Chinese (Simplified)",
        1028 => "Chinese (Traditional)",
        1025 => "Arabic",
        1037 => "Hebrew",
        1045 => "Polish",
        1053 => "Swedish",
        1035 => "Finnish",
        1030 => "Danish",
        1044 => "Norwegian",
        _ => Language > 0 ? $"LCID {Language}" : "Unknown"
    };
}

/// <summary>
/// Represents the subsites result for a single site collection.
/// </summary>
public class SiteSubsitesResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<SubsiteReportItem> Subsites { get; set; } = [];
    public int SubsiteCount => Subsites.Count;
    public bool HasSubsites => Subsites.Count > 0;
}

/// <summary>
/// Represents the complete result of a subsites report task execution.
/// </summary>
public class SubsitesReportResult
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
    public List<SiteSubsitesResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    public int SitesWithSubsites => SiteResults.Count(s => s.Success && s.HasSubsites);
    public int SitesWithoutSubsites => SiteResults.Count(s => s.Success && !s.HasSubsites);
    public int TotalSubsites => SiteResults.Where(s => s.Success).Sum(s => s.SubsiteCount);

    /// <summary>
    /// Returns all subsites flattened across all sites.
    /// </summary>
    public IEnumerable<SubsiteReportItem> GetAllSubsites() =>
        SiteResults.Where(s => s.Success).SelectMany(s => s.Subsites);

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }
}

/// <summary>
/// Export model for subsites report CSV (all subsites).
/// </summary>
public class SubsiteExportItem
{
    public string SiteCollectionUrl { get; set; } = string.Empty;
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string SubsiteUrl { get; set; } = string.Empty;
    public string SubsiteTitle { get; set; } = string.Empty;
    public string WebTemplate { get; set; } = string.Empty;
    public DateTime Created { get; set; }
    public DateTime LastModified { get; set; }
    public string Language { get; set; } = string.Empty;
}

/// <summary>
/// Export model for subsites report site summary CSV.
/// </summary>
public class SubsitesSiteSummaryExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public int SubsiteCount { get; set; }
    public string HasSubsites { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

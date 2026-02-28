namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents a single OneNote notebook folder found in a document library.
/// </summary>
public class BrokenOneNoteItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string LibraryTitle { get; set; } = string.Empty;
    public Guid LibraryId { get; set; }
    public string FolderName { get; set; } = string.Empty;
    public string FolderServerRelativeUrl { get; set; } = string.Empty;
    public int ItemId { get; set; }
    public string HtmlFileType { get; set; } = string.Empty;
    public bool IsBroken => !string.Equals(HtmlFileType, "OneNote.Notebook", StringComparison.OrdinalIgnoreCase);
    public bool IsFixed { get; set; }
    public string? FixError { get; set; }

    public string StatusDescription => IsFixed ? "Fixed" : IsBroken ? "Broken" : "Healthy";
}

/// <summary>
/// Represents broken OneNote notebook results for a single SharePoint site.
/// </summary>
public class SiteBrokenOneNoteResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<BrokenOneNoteItem> Notebooks { get; set; } = [];
    public int TotalNotebooks => Notebooks.Count;
    public int BrokenCount => Notebooks.Count(n => n.IsBroken && !n.IsFixed);
    public int FixedCount => Notebooks.Count(n => n.IsFixed);
}

/// <summary>
/// Represents the complete result of a broken OneNote notebooks report task execution.
/// </summary>
public class BrokenOneNoteReportResult
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
    public List<SiteBrokenOneNoteResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets all notebooks flattened across all sites.
    /// </summary>
    public IEnumerable<BrokenOneNoteItem> GetAllNotebooks()
    {
        foreach (var siteResult in SiteResults)
        {
            foreach (var notebook in siteResult.Notebooks)
            {
                yield return notebook;
            }
        }
    }

    /// <summary>
    /// Gets only broken notebooks (not yet fixed) flattened across all sites.
    /// </summary>
    public IEnumerable<BrokenOneNoteItem> GetBrokenNotebooks()
    {
        return GetAllNotebooks().Where(n => n.IsBroken && !n.IsFixed);
    }

    public int TotalNotebooksFound => SiteResults.Sum(s => s.TotalNotebooks);
    public int TotalBroken => SiteResults.Sum(s => s.BrokenCount);
    public int TotalFixed => SiteResults.Sum(s => s.FixedCount);

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }
}

/// <summary>
/// Export model for broken OneNote notebooks report CSV.
/// </summary>
public class BrokenOneNoteExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string LibraryTitle { get; set; } = string.Empty;
    public string FolderName { get; set; } = string.Empty;
    public string FolderPath { get; set; } = string.Empty;
    public int ItemId { get; set; }
    public string HtmlFileType { get; set; } = string.Empty;
    public string IsBroken { get; set; } = string.Empty;
    public string IsFixed { get; set; } = string.Empty;
    public string FixError { get; set; } = string.Empty;
}

/// <summary>
/// Export model for broken OneNote notebooks site summary CSV.
/// </summary>
public class BrokenOneNoteSiteSummaryExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public int TotalNotebooks { get; set; }
    public int BrokenCount { get; set; }
    public int FixedCount { get; set; }
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

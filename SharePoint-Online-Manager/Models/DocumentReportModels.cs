namespace SharePointOnlineManager.Models;

/// <summary>
/// Configuration for a document report task.
/// </summary>
public class DocumentReportConfiguration
{
    public Guid ConnectionId { get; set; }
    public List<string> TargetSiteUrls { get; set; } = [];
    public bool IncludeHiddenLibraries { get; set; } = false;
    public bool IncludeSubfolders { get; set; } = true;
    public bool IncludeVersionCount { get; set; } = true;
    public string ExtensionFilter { get; set; } = string.Empty;
}

/// <summary>
/// Represents a single document/file from a SharePoint library.
/// </summary>
public class DocumentReportItem
{
    public string FileName { get; set; } = string.Empty;
    public string Extension { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public string SizeFormatted => FormatSize(SizeBytes);
    public DateTime CreatedDate { get; set; }
    public string CreatedBy { get; set; } = string.Empty;
    public DateTime ModifiedDate { get; set; }
    public string ModifiedBy { get; set; } = string.Empty;
    public string FileUrl { get; set; } = string.Empty;
    public string ServerRelativeUrl { get; set; } = string.Empty;
    public string SiteCollectionUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string LibraryTitle { get; set; } = string.Empty;
    public int VersionCount { get; set; }
    public string FolderPath { get; set; } = string.Empty;

    private static string FormatSize(long bytes)
    {
        string[] sizes = ["B", "KB", "MB", "GB", "TB"];
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }
}

/// <summary>
/// Represents document results for a single SharePoint site.
/// </summary>
public class SiteDocumentResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<DocumentReportItem> Documents { get; set; } = [];
    public int LibrariesProcessed { get; set; }
    public int TotalDocuments => Documents.Count;
    public long TotalSizeBytes => Documents.Sum(d => d.SizeBytes);
}

/// <summary>
/// Represents the complete result of a document report task execution.
/// </summary>
public class DocumentReportResult
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
    public List<SiteDocumentResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets all documents flattened across all sites.
    /// </summary>
    public IEnumerable<DocumentReportItem> GetAllDocuments()
    {
        foreach (var siteResult in SiteResults)
        {
            foreach (var doc in siteResult.Documents)
            {
                yield return doc;
            }
        }
    }

    /// <summary>
    /// Gets summary statistics for the report.
    /// </summary>
    public (int totalDocuments, long totalSize, int totalLibraries) GetSummary()
    {
        var totalDocs = SiteResults.Sum(s => s.TotalDocuments);
        var totalSize = SiteResults.Sum(s => s.TotalSizeBytes);
        var totalLibs = SiteResults.Sum(s => s.LibrariesProcessed);
        return (totalDocs, totalSize, totalLibs);
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
/// Export model for document report with folder level columns for CSV export.
/// </summary>
public class DocumentReportExportItem
{
    public string FileName { get; set; } = string.Empty;
    public string Extension { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public string SizeFormatted { get; set; } = string.Empty;
    public DateTime CreatedDate { get; set; }
    public string CreatedBy { get; set; } = string.Empty;
    public DateTime ModifiedDate { get; set; }
    public string ModifiedBy { get; set; } = string.Empty;
    public string FileUrl { get; set; } = string.Empty;
    public string SiteCollectionUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string LibraryTitle { get; set; } = string.Empty;
    public int VersionCount { get; set; }
    public string Level1 { get; set; } = string.Empty;
    public string Level2 { get; set; } = string.Empty;
    public string Level3 { get; set; } = string.Empty;
    public string Level4 { get; set; } = string.Empty;
    public string Level5 { get; set; } = string.Empty;
    public string Level6 { get; set; } = string.Empty;
    public string Level7 { get; set; } = string.Empty;
    public string Level8 { get; set; } = string.Empty;
    public string Level9 { get; set; } = string.Empty;
    public string Level10 { get; set; } = string.Empty;

    /// <summary>
    /// Creates an export item from a document report item with folder levels populated.
    /// </summary>
    public static DocumentReportExportItem FromDocumentReportItem(DocumentReportItem item, string siteUrl)
    {
        var exportItem = new DocumentReportExportItem
        {
            FileName = item.FileName,
            Extension = item.Extension,
            SizeBytes = item.SizeBytes,
            SizeFormatted = item.SizeFormatted,
            CreatedDate = item.CreatedDate,
            CreatedBy = item.CreatedBy,
            ModifiedDate = item.ModifiedDate,
            ModifiedBy = item.ModifiedBy,
            FileUrl = item.FileUrl,
            SiteCollectionUrl = item.SiteCollectionUrl,
            SiteTitle = item.SiteTitle,
            LibraryTitle = item.LibraryTitle,
            VersionCount = item.VersionCount
        };

        // Extract folder levels from the folder path
        var folderPath = item.FolderPath;
        if (!string.IsNullOrEmpty(folderPath))
        {
            // Remove leading slash if present
            folderPath = folderPath.TrimStart('/');
            var segments = folderPath.Split('/', StringSplitOptions.RemoveEmptyEntries);

            // Build cumulative URLs for each level
            var baseUrl = siteUrl.TrimEnd('/');
            var cumulativePath = "";

            for (int i = 0; i < segments.Length && i < 10; i++)
            {
                cumulativePath += "/" + segments[i];
                var levelUrl = baseUrl + cumulativePath;

                switch (i)
                {
                    case 0: exportItem.Level1 = levelUrl; break;
                    case 1: exportItem.Level2 = levelUrl; break;
                    case 2: exportItem.Level3 = levelUrl; break;
                    case 3: exportItem.Level4 = levelUrl; break;
                    case 4: exportItem.Level5 = levelUrl; break;
                    case 5: exportItem.Level6 = levelUrl; break;
                    case 6: exportItem.Level7 = levelUrl; break;
                    case 7: exportItem.Level8 = levelUrl; break;
                    case 8: exportItem.Level9 = levelUrl; break;
                    case 9: exportItem.Level10 = levelUrl; break;
                }
            }
        }

        return exportItem;
    }
}

/// <summary>
/// Export model for document report site summary.
/// </summary>
public class DocumentReportSiteSummaryExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public int LibrariesProcessed { get; set; }
    public int TotalDocuments { get; set; }
    public long TotalSizeBytes { get; set; }
    public string TotalSizeFormatted { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

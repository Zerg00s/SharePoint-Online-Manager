namespace SharePointOnlineManager.Models;

/// <summary>
/// Defines the item type for document comparison.
/// </summary>
public enum DocumentCompareItemType
{
    File,
    Folder
}

/// <summary>
/// Defines the comparison status for a document.
/// </summary>
public enum DocumentCompareStatus
{
    /// <summary>
    /// Document exists on both sides (regardless of size difference).
    /// </summary>
    Found,

    /// <summary>
    /// Document exists on both sides but has a concerning size issue:
    /// - Target is 0 bytes when source > 0
    /// - Source > 50KB and target is less than 30% of source size
    /// </summary>
    SizeIssue,

    /// <summary>
    /// Document exists only on source site (not migrated).
    /// </summary>
    SourceOnly,

    /// <summary>
    /// Document exists only on target site (extra file).
    /// </summary>
    TargetOnly
}

/// <summary>
/// Configuration for a document compare task.
/// </summary>
public class DocumentCompareConfiguration
{
    public Guid SourceConnectionId { get; set; }
    public Guid TargetConnectionId { get; set; }
    public List<SiteComparePair> SitePairs { get; set; } = [];
    public List<string> ExcludedLibraries { get; set; } = [];
    public bool IncludeHiddenLibraries { get; set; } = false;
    public bool IncludeAspxPages { get; set; } = false;

    /// <summary>
    /// When true, attempts to match files by normalizing special characters
    /// that ShareGate often replaces with underscore during migration.
    /// Characters: " * : &lt; &gt; ? \ &amp; # % { } ~
    /// </summary>
    public bool UseShareGateNormalization { get; set; } = false;

    /// <summary>
    /// When true, uses cached document lists if available and not expired.
    /// Cache is valid for the duration specified in CacheExpirationHours.
    /// </summary>
    public bool UseCache { get; set; } = false;

    /// <summary>
    /// Number of hours before cached document lists expire. Default is 48 hours (2 days).
    /// </summary>
    public int CacheExpirationHours { get; set; } = 48;

    /// <summary>
    /// Gets the default list of libraries that should always be excluded.
    /// </summary>
    public static List<string> DefaultExcludedLibraries =>
    [
        "Style Library",
        "Form Templates",
        "Site Collection Documents",
        "Site Collection Images",
        "_catalogs/hubsite",
        "Preservation Hold Library",
        "appdata"
    ];
}

/// <summary>
/// Represents a source document item for comparison.
/// </summary>
public class DocumentCompareSourceItem
{
    public int Id { get; set; }
    public string FileName { get; set; } = string.Empty;
    public string ServerRelativeUrl { get; set; } = string.Empty;
    public string RelativePath { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public int VersionCount { get; set; }
    public string LibraryTitle { get; set; } = string.Empty;
    public DocumentCompareItemType ItemType { get; set; } = DocumentCompareItemType.File;
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }
}

/// <summary>
/// Represents a single document comparison result.
/// </summary>
public class DocumentCompareItem
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string LibraryName { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string FileExtension { get; set; } = string.Empty;
    public string RelativePath { get; set; } = string.Empty;
    public DocumentCompareItemType ItemType { get; set; } = DocumentCompareItemType.File;
    public int SourceItemId { get; set; }
    public int TargetItemId { get; set; }
    public long SourceSizeBytes { get; set; }
    public long TargetSizeBytes { get; set; }
    public int SourceVersionCount { get; set; }
    public int TargetVersionCount { get; set; }
    public string SourceAbsolutePath { get; set; } = string.Empty;
    public string TargetAbsolutePath { get; set; } = string.Empty;
    public DateTime? SourceCreated { get; set; }
    public DateTime? SourceModified { get; set; }
    public DateTime? TargetCreated { get; set; }
    public DateTime? TargetModified { get; set; }
    public DocumentCompareStatus Status { get; set; }

    /// <summary>
    /// Gets the size difference as a percentage.
    /// </summary>
    public double SizeDifferencePercent
    {
        get
        {
            if (SourceSizeBytes == 0)
                return TargetSizeBytes == 0 ? 0 : 100;
            return Math.Abs((double)(TargetSizeBytes - SourceSizeBytes) / SourceSizeBytes * 100);
        }
    }

    /// <summary>
    /// Gets a display string for the status.
    /// </summary>
    public string StatusDescription => Status switch
    {
        DocumentCompareStatus.Found => "Found",
        DocumentCompareStatus.SizeIssue => "Size Issue",
        DocumentCompareStatus.SourceOnly => "Source Only",
        DocumentCompareStatus.TargetOnly => "Target Only",
        _ => Status.ToString()
    };

    /// <summary>
    /// Indicates if this document has a size issue (0 bytes on target, or significant shrinkage).
    /// </summary>
    public bool HasSizeIssue => Status == DocumentCompareStatus.SizeIssue;

    /// <summary>
    /// Indicates if the source copy was modified more than 24 hours after the target copy,
    /// meaning the target is stale and needs re-migration.
    /// </summary>
    public bool IsNewerAtSource =>
        SourceModified.HasValue && TargetModified.HasValue &&
        (Status == DocumentCompareStatus.Found || Status == DocumentCompareStatus.SizeIssue) &&
        (SourceModified.Value - TargetModified.Value).TotalHours > 24;
}

/// <summary>
/// Represents comparison results for a single site pair.
/// </summary>
public class SiteDocumentCompareResult
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string SourceSiteTitle { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string TargetSiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<DocumentCompareItem> DocumentComparisons { get; set; } = [];
    public int LibrariesProcessed { get; set; }

    /// <summary>
    /// Gets the count of documents found on both sides (regardless of size).
    /// </summary>
    public int FoundCount => DocumentComparisons.Count(d => d.Status == DocumentCompareStatus.Found || d.Status == DocumentCompareStatus.SizeIssue);

    /// <summary>
    /// Gets the count of documents with concerning size issues.
    /// </summary>
    public int SizeIssueCount => DocumentComparisons.Count(d => d.Status == DocumentCompareStatus.SizeIssue);

    /// <summary>
    /// Gets the count of documents that exist only on source.
    /// </summary>
    public int SourceOnlyCount => DocumentComparisons.Count(d => d.Status == DocumentCompareStatus.SourceOnly);

    /// <summary>
    /// Gets the count of documents that exist only on target.
    /// </summary>
    public int TargetOnlyCount => DocumentComparisons.Count(d => d.Status == DocumentCompareStatus.TargetOnly);

    /// <summary>
    /// Gets the count of documents where source is newer than target by more than 24 hours (stale target).
    /// </summary>
    public int NewerAtSourceCount => DocumentComparisons.Count(d => d.IsNewerAtSource);

    /// <summary>
    /// Gets the total document count.
    /// </summary>
    public int TotalDocuments => DocumentComparisons.Count;

    /// <summary>
    /// Gets the total number of unique documents on source (found + source only).
    /// </summary>
    public int TotalSourceDocuments => FoundCount + SourceOnlyCount;

    /// <summary>
    /// Gets the total number of unique documents on target (found + target only).
    /// </summary>
    public int TotalTargetDocuments => FoundCount + TargetOnlyCount;

    /// <summary>
    /// Gets the total size of documents on source in bytes.
    /// </summary>
    public long TotalSourceSizeBytes => DocumentComparisons
        .Where(d => d.Status != DocumentCompareStatus.TargetOnly)
        .Sum(d => d.SourceSizeBytes);

    /// <summary>
    /// Gets the total size of documents on target in bytes.
    /// </summary>
    public long TotalTargetSizeBytes => DocumentComparisons
        .Where(d => d.Status != DocumentCompareStatus.SourceOnly)
        .Sum(d => d.TargetSizeBytes);

    /// <summary>
    /// Gets the percentage of source documents found on target.
    /// </summary>
    public double PercentFound => TotalSourceDocuments > 0
        ? (double)FoundCount / TotalSourceDocuments * 100
        : 100;

    /// <summary>
    /// Gets the percentage of source documents not found on target.
    /// </summary>
    public double PercentNotFound => TotalSourceDocuments > 0
        ? (double)SourceOnlyCount / TotalSourceDocuments * 100
        : 0;

    /// <summary>
    /// Gets the percentage of target documents that are extra (not on source).
    /// </summary>
    public double PercentTargetOnly => TotalTargetDocuments > 0
        ? (double)TargetOnlyCount / TotalTargetDocuments * 100
        : 0;

    /// <summary>
    /// Gets the average version count for source documents.
    /// </summary>
    public double AvgSourceVersions
    {
        get
        {
            var sourceDocs = DocumentComparisons.Where(d => d.Status != DocumentCompareStatus.TargetOnly).ToList();
            return sourceDocs.Count > 0 ? sourceDocs.Average(d => d.SourceVersionCount) : 0;
        }
    }

    /// <summary>
    /// Gets the average version count for target documents.
    /// </summary>
    public double AvgTargetVersions
    {
        get
        {
            var targetDocs = DocumentComparisons.Where(d => d.Status != DocumentCompareStatus.SourceOnly).ToList();
            return targetDocs.Count > 0 ? targetDocs.Average(d => d.TargetVersionCount) : 0;
        }
    }

    /// <summary>
    /// Indicates whether this site pair has any issues.
    /// </summary>
    public bool HasIssues => !Success || SizeIssueCount > 0 || SourceOnlyCount > 0 || NewerAtSourceCount > 0;
}

/// <summary>
/// Represents the complete result of a document compare task execution.
/// </summary>
public class DocumentCompareResult
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
    public int ThrottleRetryCount { get; set; }
    public List<SiteDocumentCompareResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets all document comparison items flattened across all site pairs.
    /// </summary>
    public IEnumerable<DocumentCompareItem> GetAllDocumentComparisons()
    {
        foreach (var siteResult in SiteResults)
        {
            foreach (var comparison in siteResult.DocumentComparisons)
            {
                yield return comparison;
            }
        }
    }

    /// <summary>
    /// Gets site results that have issues (errors, mismatches, or missing documents).
    /// </summary>
    public IEnumerable<SiteDocumentCompareResult> GetSitesWithIssues()
    {
        return SiteResults.Where(s => s.HasIssues);
    }

    /// <summary>
    /// Gets the summary statistics.
    /// </summary>
    public (int TotalFound, int TotalSizeIssues, int TotalSourceOnly, int TotalTargetOnly, int TotalNewerAtSource) GetSummary()
    {
        int found = 0, sizeIssues = 0, sourceOnly = 0, targetOnly = 0, newerAtSource = 0;
        foreach (var site in SiteResults)
        {
            found += site.FoundCount;
            sizeIssues += site.SizeIssueCount;
            sourceOnly += site.SourceOnlyCount;
            targetOnly += site.TargetOnlyCount;
            newerAtSource += site.NewerAtSourceCount;
        }
        return (found, sizeIssues, sourceOnly, targetOnly, newerAtSource);
    }

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }

    /// <summary>
    /// Gets total source size in bytes across all sites.
    /// </summary>
    public long TotalSourceSizeBytes => SiteResults.Sum(s => s.TotalSourceSizeBytes);

    /// <summary>
    /// Gets total target size in bytes across all sites.
    /// </summary>
    public long TotalTargetSizeBytes => SiteResults.Sum(s => s.TotalTargetSizeBytes);

    /// <summary>
    /// Gets the overall average source version count.
    /// </summary>
    public double OverallAvgSourceVersions
    {
        get
        {
            var allDocs = GetAllDocumentComparisons()
                .Where(d => d.Status != DocumentCompareStatus.TargetOnly)
                .ToList();
            return allDocs.Count > 0 ? allDocs.Average(d => d.SourceVersionCount) : 0;
        }
    }

    /// <summary>
    /// Gets the overall average target version count.
    /// </summary>
    public double OverallAvgTargetVersions
    {
        get
        {
            var allDocs = GetAllDocumentComparisons()
                .Where(d => d.Status != DocumentCompareStatus.SourceOnly)
                .ToList();
            return allDocs.Count > 0 ? allDocs.Average(d => d.TargetVersionCount) : 0;
        }
    }

    /// <summary>
    /// Gets overall migration completeness percentage.
    /// </summary>
    public double MigrationCompletenessPercent
    {
        get
        {
            var (found, _, sourceOnly, _, _) = GetSummary();
            var totalSource = found + sourceOnly;
            return totalSource > 0 ? (double)found / totalSource * 100 : 100;
        }
    }
}

/// <summary>
/// Export model for document comparison results (CSV).
/// </summary>
public class DocumentCompareExportItem
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string LibraryName { get; set; } = string.Empty;
    public string ItemType { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string FileExtension { get; set; } = string.Empty;
    public int SourceItemId { get; set; }
    public int TargetItemId { get; set; }
    public long SourceSizeBytes { get; set; }
    public long TargetSizeBytes { get; set; }
    public string SizeDifferencePercent { get; set; } = string.Empty;
    public int SourceVersionCount { get; set; }
    public int TargetVersionCount { get; set; }
    public string SourceAbsolutePath { get; set; } = string.Empty;
    public string TargetAbsolutePath { get; set; } = string.Empty;
    public DateTime? SourceCreated { get; set; }
    public DateTime? SourceModified { get; set; }
    public DateTime? TargetCreated { get; set; }
    public DateTime? TargetModified { get; set; }
    public string Status { get; set; } = string.Empty;
    public string Note { get; set; } = string.Empty;
}

/// <summary>
/// Export model for site summary in XLSX.
/// </summary>
public class DocumentCompareSiteSummary
{
    public string SourceSiteUrl { get; set; } = string.Empty;
    public string SourceSiteTitle { get; set; } = string.Empty;
    public string TargetSiteUrl { get; set; } = string.Empty;
    public string TargetSiteTitle { get; set; } = string.Empty;
    public int TotalSourceDocuments { get; set; }
    public int TotalTargetDocuments { get; set; }
    public int Found { get; set; }
    public int SizeIssues { get; set; }
    public int SourceOnly { get; set; }
    public int TargetOnly { get; set; }
    public double PercentFound { get; set; }
    public double PercentNotFound { get; set; }
    public double PercentTargetOnly { get; set; }
    public long SourceSizeBytes { get; set; }
    public long TargetSizeBytes { get; set; }
    public double AvgSourceVersions { get; set; }
    public double AvgTargetVersions { get; set; }
    public int LibrariesProcessed { get; set; }
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

/// <summary>
/// Cache entry for document library scan results.
/// </summary>
public class DocumentCompareCacheEntry
{
    public DateTime CachedAt { get; set; } = DateTime.UtcNow;
    public string SiteUrl { get; set; } = string.Empty;
    public string LibraryTitle { get; set; } = string.Empty;
    public List<DocumentCompareSourceItem> Documents { get; set; } = [];

    /// <summary>
    /// Checks if this cache entry is still valid.
    /// </summary>
    public bool IsValid(int expirationHours)
    {
        return DateTime.UtcNow - CachedAt < TimeSpan.FromHours(expirationHours);
    }
}

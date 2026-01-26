using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Service for exporting data to CSV format.
/// </summary>
public class CsvExporter
{
    /// <summary>
    /// Exports task results to a CSV file.
    /// </summary>
    public void ExportListReport(TaskResult result, string filePath, IEnumerable<string>? excludedLists = null)
    {
        var excludeSet = excludedLists != null
            ? new HashSet<string>(excludedLists, StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var items = result.GetAllListItems()
            .Where(item => !excludeSet.Contains(item.ListTitle))
            .Select(item => new ListReportExportItem
            {
                SiteUrl = item.SiteUrl,
                SiteTitle = item.SiteTitle,
                ListTitle = item.ListTitle,
                ListUrl = item.ListUrl,
                ListType = item.ListType,
                ItemCount = item.ItemCount,
                Hidden = item.Hidden ? "Yes" : "No",
                Created = item.Created,
                LastModified = item.LastModified
            })
            .ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports site collections to a CSV file.
    /// </summary>
    public void ExportSiteCollections(List<SiteCollection> sites, string filePath)
    {
        var items = sites.Select(s => new SiteExportItem
        {
            Url = s.Url,
            Title = s.Title,
            SiteType = s.SiteTypeDescription,
            Template = s.Template,
            Owner = s.Owner,
            StorageUsed = s.StorageUsedFormatted,
            Status = s.Status,
            LastModified = s.LastModifiedDate
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports list compare results to a CSV file.
    /// </summary>
    public void ExportListCompareReport(ListCompareResult result, string filePath)
    {
        var items = result.GetAllListComparisons().Select(c => new ListCompareExportItem
        {
            SourceSiteUrl = c.SourceSiteUrl,
            TargetSiteUrl = c.TargetSiteUrl,
            ListTitle = c.ListTitle,
            ListType = c.ListType,
            SourceCount = c.SourceCount,
            TargetCount = c.TargetCount,
            Difference = c.Difference,
            PercentDifference = $"{c.PercentDifference:F1}%",
            Status = c.StatusDescription
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports sites with issues to a CSV file.
    /// </summary>
    public void ExportIssuesSummary(ListCompareResult result, string filePath)
    {
        var items = result.GetSitesWithIssues().Select(s => new SiteIssueExportItem
        {
            SourceSiteUrl = s.SourceSiteUrl,
            TargetSiteUrl = s.TargetSiteUrl,
            Mismatches = s.MismatchCount,
            SourceOnly = s.SourceOnlyCount,
            TargetOnly = s.TargetOnlyCount,
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports list mapping with full URLs, filtering out system lists but keeping Site Pages.
    /// </summary>
    public void ExportListMapping(ListCompareResult result, string filePath)
    {
        var excludedLists = new HashSet<string>(
            ListCompareConfiguration.DefaultExcludedLists,
            StringComparer.OrdinalIgnoreCase);

        var neverExcluded = new HashSet<string>(
            ListCompareConfiguration.NeverExcludedLists,
            StringComparer.OrdinalIgnoreCase);

        var items = result.GetAllListComparisons()
            .Where(c => neverExcluded.Contains(c.ListTitle) || !excludedLists.Contains(c.ListTitle))
            .Select(c => new ListMappingExportItem
            {
                SourceSiteUrl = c.SourceSiteUrl,
                SourceSiteTitle = c.SourceSiteTitle,
                TargetSiteUrl = c.TargetSiteUrl,
                TargetSiteTitle = c.TargetSiteTitle,
                ListTitle = c.ListTitle,
                ListType = c.ListType,
                SourceListUrl = c.SourceListUrl,
                TargetListUrl = c.TargetListUrl,
                SourceCount = c.SourceCount,
                TargetCount = c.TargetCount,
                Status = c.StatusDescription
            }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports document report results to a CSV file with folder level columns.
    /// </summary>
    public void ExportDocumentReport(DocumentReportResult result, string filePath)
    {
        var items = result.GetAllDocuments()
            .Select(d => DocumentReportExportItem.FromDocumentReportItem(d, d.SiteCollectionUrl))
            .ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports document report site summary to a CSV file.
    /// </summary>
    public void ExportDocumentReportSummary(DocumentReportResult result, string filePath)
    {
        var items = result.SiteResults.Select(s => new DocumentReportSiteSummaryExportItem
        {
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            LibrariesProcessed = s.LibrariesProcessed,
            TotalDocuments = s.TotalDocuments,
            TotalSizeBytes = s.TotalSizeBytes,
            TotalSizeFormatted = FormatSize(s.TotalSizeBytes),
            Status = s.Success ? "Success" : "Failed",
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports permission report results to a CSV file.
    /// </summary>
    public void ExportPermissionReport(PermissionReportResult result, string filePath)
    {
        var items = result.GetAllPermissions()
            .Select(p => new PermissionReportExportItem
            {
                ObjectType = p.ObjectTypeDescription,
                SiteCollectionUrl = p.SiteCollectionUrl,
                SiteUrl = p.SiteUrl,
                SiteTitle = p.SiteTitle,
                ObjectTitle = p.ObjectTitle,
                ObjectUrl = p.ObjectUrl,
                ObjectPath = p.ObjectPath,
                PrincipalName = p.PrincipalName,
                PrincipalType = p.PrincipalType,
                PrincipalLogin = p.PrincipalLogin,
                PermissionLevel = p.PermissionLevel,
                IsInherited = p.IsInherited ? "Yes" : "No",
                InheritedFrom = p.InheritedFrom
            })
            .ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports permission report site summary to a CSV file.
    /// </summary>
    public void ExportPermissionReportSummary(PermissionReportResult result, string filePath)
    {
        var items = result.SiteResults.Select(s => new PermissionReportSiteSummaryExportItem
        {
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            TotalPermissions = s.TotalPermissions,
            UniqueObjects = s.UniquePermissionObjects,
            Status = s.Success ? "Success" : "Failed",
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

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

    private static void ExportToCsv<T>(List<T> items, string filePath)
    {
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true
        };

        using var writer = new StreamWriter(filePath);
        using var csv = new CsvWriter(writer, config);
        csv.WriteRecords(items);
    }
}

/// <summary>
/// Export model for site collections.
/// </summary>
public class SiteExportItem
{
    public string Url { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string SiteType { get; set; } = string.Empty;
    public string Template { get; set; } = string.Empty;
    public string Owner { get; set; } = string.Empty;
    public string StorageUsed { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public DateTime LastModified { get; set; }
}

/// <summary>
/// Export model for list report results.
/// </summary>
public class ListReportExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string ListTitle { get; set; } = string.Empty;
    public string ListUrl { get; set; } = string.Empty;
    public string ListType { get; set; } = string.Empty;
    public int ItemCount { get; set; }
    public string Hidden { get; set; } = string.Empty;
    public DateTime Created { get; set; }
    public DateTime LastModified { get; set; }
}

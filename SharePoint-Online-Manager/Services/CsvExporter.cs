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
            TemplateDescription = s.TemplateDescription,
            Owner = s.Owner,
            Standalone = !s.IsGroupConnected && s.ChannelType == 0 ? "Yes" : "No",
            Group = s.IsGroupConnected ? "Yes" : "No",
            Channel = s.ChannelTypeDisplay,
            Hub = s.HubDisplay,
            StorageUsedGB = Math.Round(s.StorageUsed / 1073741824.0, 2),
            FileCount = s.FileCount,
            PageViews = s.PageViews,
            LastActivity = s.LastActivityDisplay,
            ExternalSharing = s.ExternalSharing,
            State = s.StateDisplay,
            Language = s.LanguageDisplay,
            IsDeleted = s.IsDeleted ? "Yes" : "No",
            TimeDeleted = s.TimeDeleted
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
    /// Exports a collection of list compare items to a CSV file.
    /// </summary>
    public void ExportListCompareItems(IEnumerable<ListCompareItem> items, string filePath)
    {
        var exportItems = items.Select(c => new ListCompareExportItem
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

        ExportToCsv(exportItems, filePath);
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

    /// <summary>
    /// Exports navigation settings comparison results to a CSV file.
    /// </summary>
    public void ExportNavigationSettingsReport(NavigationSettingsResult result, string filePath)
    {
        var items = result.SiteResults.Select(s => new NavigationSettingsExportItem
        {
            SourceSiteUrl = s.SourceSiteUrl,
            SourceSiteTitle = s.SourceSiteTitle,
            TargetSiteUrl = s.TargetSiteUrl,
            TargetSiteTitle = s.TargetSiteTitle,
            SourceHorizontalNav = s.SourceHorizontalQuickLaunch ? "Yes" : "No",
            TargetHorizontalNav = s.TargetHorizontalQuickLaunch ? "Yes" : "No",
            HorizontalNavMatch = s.HorizontalQuickLaunchMatches ? "Yes" : "No",
            SourceMegaMenu = s.SourceMegaMenuEnabled ? "Yes" : "No",
            TargetMegaMenu = s.TargetMegaMenuEnabled ? "Yes" : "No",
            MegaMenuMatch = s.MegaMenuEnabledMatches ? "Yes" : "No",
            Status = s.StatusDescription,
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports document comparison results to a CSV file.
    /// </summary>
    public void ExportDocumentCompareReport(DocumentCompareResult result, string filePath)
    {
        var items = result.GetAllDocumentComparisons().Select(d => ToDocumentCompareExportItem(d)).ToList();
        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports a collection of document compare items to a CSV file.
    /// </summary>
    public void ExportDocumentCompareItems(IEnumerable<DocumentCompareItem> items, string filePath)
    {
        var exportItems = items.Select(d => ToDocumentCompareExportItem(d)).ToList();
        ExportToCsv(exportItems, filePath);
    }

    /// <summary>
    /// Exports a single site's document comparison results to a CSV file.
    /// </summary>
    public void ExportSiteDocumentCompareReport(SiteDocumentCompareResult siteResult, string filePath)
    {
        var items = siteResult.DocumentComparisons.Select(d => ToDocumentCompareExportItem(d)).ToList();
        ExportToCsv(items, filePath);
    }

    private static DocumentCompareExportItem ToDocumentCompareExportItem(DocumentCompareItem d)
    {
        var item = new DocumentCompareExportItem
        {
            SourceSiteUrl = d.SourceSiteUrl,
            TargetSiteUrl = d.TargetSiteUrl,
            LibraryName = d.LibraryName,
            ItemType = d.ItemType.ToString(),
            FileName = d.FileName,
            FileExtension = d.FileExtension,
            SourceItemId = d.SourceItemId,
            TargetItemId = d.TargetItemId,
            SourceSizeBytes = d.SourceSizeBytes,
            TargetSizeBytes = d.TargetSizeBytes,
            SizeDifferencePercent = $"{d.SizeDifferencePercent:F1}%",
            SourceVersionCount = d.SourceVersionCount,
            TargetVersionCount = d.TargetVersionCount,
            SourceAbsolutePath = d.SourceAbsolutePath,
            TargetAbsolutePath = d.TargetAbsolutePath,
            SourceCreated = d.SourceCreated,
            SourceModified = d.SourceModified,
            TargetCreated = d.TargetCreated,
            TargetModified = d.TargetModified,
            Status = d.StatusDescription
        };

        var notes = new List<string>();

        // For SourceOnly items, check if the expected target URL would exceed 400 characters
        if (d.Status == DocumentCompareStatus.SourceOnly &&
            !string.IsNullOrEmpty(d.SourceAbsolutePath) &&
            !string.IsNullOrEmpty(d.SourceSiteUrl) &&
            !string.IsNullOrEmpty(d.TargetSiteUrl))
        {
            // Expected target path = target site URL + the path portion after source site URL
            var sourceRelative = d.SourceAbsolutePath.Length > d.SourceSiteUrl.Length
                ? d.SourceAbsolutePath[d.SourceSiteUrl.Length..]
                : "";
            var expectedTargetUrl = d.TargetSiteUrl.TrimEnd('/') + sourceRelative;
            if (expectedTargetUrl.Length > 400)
            {
                notes.Add($"Target path would exceed 400 chars ({expectedTargetUrl.Length})");
            }
        }

        if (d.IsNewerAtSource)
        {
            notes.Add("Source modified >24h after target");
        }

        item.Note = string.Join("; ", notes);

        return item;
    }

    /// <summary>
    /// Exports customized lists report results to a CSV file.
    /// </summary>
    public void ExportCustomizedListsReport(CustomizedListsReportResult result, string filePath, bool customizedOnly = false)
    {
        var lists = customizedOnly ? result.GetCustomizedLists() : result.GetAllLists();
        var items = lists.Select(l => new CustomizedListsExportItem
        {
            SiteUrl = l.SiteUrl,
            SiteTitle = l.SiteTitle,
            ListTitle = l.ListTitle,
            ListType = l.ListType,
            FormType = l.FormTypeDescription,
            ItemCount = l.ItemCount,
            ListUrl = l.ListUrl,
            DefaultNewFormUrl = l.DefaultNewFormUrl,
            DefaultEditFormUrl = l.DefaultEditFormUrl,
            SpfxNewFormComponentId = l.SpfxNewFormComponentId,
            SpfxEditFormComponentId = l.SpfxEditFormComponentId,
            SpfxDisplayFormComponentId = l.SpfxDisplayFormComponentId
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports customized lists report site summary to a CSV file.
    /// </summary>
    public void ExportCustomizedListsReportSummary(CustomizedListsReportResult result, string filePath)
    {
        var items = result.SiteResults.Select(s => new CustomizedListsSiteSummaryExportItem
        {
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            TotalLists = s.TotalLists,
            CustomizedCount = s.CustomizedCount,
            PowerAppsCount = s.PowerAppsCount,
            SpfxCount = s.SpfxCount,
            Status = s.Success ? "Success" : "Failed",
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports ad hoc users report results to a CSV file.
    /// </summary>
    public void ExportAdHocUsersReport(AdHocUsersReportResult result, string filePath)
    {
        var items = result.GetAllUsers().Select(u => new AdHocUsersExportItem
        {
            SiteUrl = u.SiteUrl,
            SiteTitle = u.SiteTitle,
            LoginName = u.LoginName,
            Title = u.Title,
            Email = u.Email,
            Id = u.Id,
            IsSiteAdmin = u.IsSiteAdmin ? "Yes" : "No",
            PrincipalType = u.PrincipalType
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports ad hoc users report site summary to a CSV file.
    /// </summary>
    public void ExportAdHocUsersReportSummary(AdHocUsersReportResult result, string filePath)
    {
        var items = result.SiteResults.Select(s => new AdHocUsersSiteSummaryExportItem
        {
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            GuestCount = s.GuestCount,
            Status = s.Success ? "Success" : "Failed",
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports custom fields report results to a CSV file.
    /// </summary>
    public void ExportCustomFieldsReport(CustomFieldsReportResult result, string filePath)
    {
        var items = result.GetAllFields().Select(f => new CustomFieldsExportItem
        {
            SiteCollectionUrl = f.SiteCollectionUrl,
            SiteUrl = f.SiteUrl,
            SiteTitle = f.SiteTitle,
            ListTitle = f.ListTitle,
            ListUrl = f.ListUrl,
            ItemCount = f.ItemCount,
            ColumnName = f.ColumnName,
            InternalName = f.InternalName,
            FieldType = f.FieldType,
            Group = f.Group,
            ListCreated = f.ListCreated,
            ListModified = f.ListModified
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports custom fields report site summary to a CSV file.
    /// </summary>
    public void ExportCustomFieldsReportSummary(CustomFieldsReportResult result, string filePath)
    {
        var items = result.SiteResults.Select(s => new CustomFieldsSiteSummaryExportItem
        {
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            ListsScanned = s.ListsScanned,
            ListsWithCustomFields = s.ListsWithCustomFields,
            CustomFieldCount = s.CustomFieldCount,
            Status = s.Success ? "Success" : "Failed",
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports all subsites from the subsites report to a CSV file.
    /// </summary>
    public void ExportSubsitesReport(SubsitesReportResult result, string filePath)
    {
        var items = result.GetAllSubsites().Select(s => new SubsiteExportItem
        {
            SiteCollectionUrl = s.SiteCollectionUrl,
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            SubsiteUrl = s.SubsiteUrl,
            SubsiteTitle = s.SubsiteTitle,
            WebTemplate = s.WebTemplate,
            Created = s.Created,
            LastModified = s.LastModified,
            Language = s.LanguageDisplay
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports subsites report site summary to a CSV file (sites with subsite counts).
    /// </summary>
    public void ExportSubsitesReportSummary(SubsitesReportResult result, string filePath, bool withSubsitesOnly = false)
    {
        var sites = withSubsitesOnly
            ? result.SiteResults.Where(s => s.Success && s.HasSubsites)
            : result.SiteResults.AsEnumerable();

        var items = sites.Select(s => new SubsitesSiteSummaryExportItem
        {
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            SubsiteCount = s.SubsiteCount,
            HasSubsites = s.HasSubsites ? "Yes" : "No",
            Status = s.Success ? "Success" : "Failed",
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports publishing sites report results to a CSV file.
    /// </summary>
    public void ExportPublishingSitesReport(PublishingSitesReportResult result, string filePath, bool publishingOnly = false)
    {
        var sites = publishingOnly
            ? result.SiteResults.Where(s => s.Success && s.HasPublishing)
            : result.SiteResults.AsEnumerable();

        var items = sites.Select(s => new PublishingSitesExportItem
        {
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            PublishingInfrastructure = s.HasPublishingInfrastructure ? "Yes" : "No",
            PublishingWeb = s.HasPublishingWeb ? "Yes" : "No",
            PublishingStatus = s.PublishingStatus,
            Status = s.Success ? "Success" : "Failed",
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports publishing sites report site summary to a CSV file.
    /// </summary>
    public void ExportPublishingSitesReportSummary(PublishingSitesReportResult result, string filePath)
    {
        var items = result.SiteResults.Select(s => new PublishingSitesExportItem
        {
            SiteUrl = s.SiteUrl,
            SiteTitle = s.SiteTitle,
            PublishingInfrastructure = s.HasPublishingInfrastructure ? "Yes" : "No",
            PublishingWeb = s.HasPublishingWeb ? "Yes" : "No",
            PublishingStatus = s.PublishingStatus,
            Status = s.Success ? "Success" : "Failed",
            Error = s.ErrorMessage ?? ""
        }).ToList();

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports site access check results to a CSV file.
    /// </summary>
    public void ExportSiteAccessReport(SiteAccessResult result, string filePath, bool sourceOnly = false, bool targetOnly = false, bool issuesOnly = false)
    {
        var items = new List<SiteAccessExportItem>();

        foreach (var pair in result.PairResults)
        {
            // Add source site
            if (!targetOnly && (!issuesOnly || pair.SourceResult.HasIssue))
            {
                items.Add(new SiteAccessExportItem
                {
                    SiteUrl = pair.SourceResult.SiteUrl,
                    SiteTitle = pair.SourceResult.SiteTitle,
                    Side = "Source",
                    Status = pair.SourceResult.StatusDescription,
                    AccountUsed = pair.SourceResult.AccountUsed ?? "",
                    ErrorMessage = pair.SourceResult.ErrorMessage ?? ""
                });
            }

            // Add target site
            if (!sourceOnly && (!issuesOnly || pair.TargetResult.HasIssue))
            {
                items.Add(new SiteAccessExportItem
                {
                    SiteUrl = pair.TargetResult.SiteUrl,
                    SiteTitle = pair.TargetResult.SiteTitle,
                    Side = "Target",
                    Status = pair.TargetResult.StatusDescription,
                    AccountUsed = pair.TargetResult.AccountUsed ?? "",
                    ErrorMessage = pair.TargetResult.ErrorMessage ?? ""
                });
            }
        }

        ExportToCsv(items, filePath);
    }

    /// <summary>
    /// Exports only sites with access issues to a CSV file.
    /// </summary>
    public void ExportSiteAccessIssues(SiteAccessResult result, string filePath, bool sourceIssuesOnly = false, bool targetIssuesOnly = false)
    {
        var items = new List<SiteAccessExportItem>();

        foreach (var pair in result.PairResults)
        {
            // Add source issues
            if (!targetIssuesOnly && pair.SourceResult.HasIssue)
            {
                items.Add(new SiteAccessExportItem
                {
                    SiteUrl = pair.SourceResult.SiteUrl,
                    SiteTitle = pair.SourceResult.SiteTitle,
                    Side = "Source",
                    Status = pair.SourceResult.StatusDescription,
                    AccountUsed = pair.SourceResult.AccountUsed ?? "",
                    ErrorMessage = pair.SourceResult.ErrorMessage ?? ""
                });
            }

            // Add target issues
            if (!sourceIssuesOnly && pair.TargetResult.HasIssue)
            {
                items.Add(new SiteAccessExportItem
                {
                    SiteUrl = pair.TargetResult.SiteUrl,
                    SiteTitle = pair.TargetResult.SiteTitle,
                    Side = "Target",
                    Status = pair.TargetResult.StatusDescription,
                    AccountUsed = pair.TargetResult.AccountUsed ?? "",
                    ErrorMessage = pair.TargetResult.ErrorMessage ?? ""
                });
            }
        }

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
    public string TemplateDescription { get; set; } = string.Empty;
    public string Owner { get; set; } = string.Empty;
    public string Standalone { get; set; } = string.Empty;
    public string Group { get; set; } = string.Empty;
    public string Channel { get; set; } = string.Empty;
    public string Hub { get; set; } = string.Empty;
    public double StorageUsedGB { get; set; }
    public int FileCount { get; set; }
    public int PageViews { get; set; }
    public string LastActivity { get; set; } = string.Empty;
    public string ExternalSharing { get; set; } = string.Empty;
    public string State { get; set; } = string.Empty;
    public string Language { get; set; } = string.Empty;
    public string IsDeleted { get; set; } = string.Empty;
    public DateTime? TimeDeleted { get; set; }
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

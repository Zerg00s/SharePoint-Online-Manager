namespace SharePointOnlineManager.Models;

/// <summary>
/// Defines the form customization type for a SharePoint list.
/// </summary>
public enum ListFormType
{
    Default,
    PowerApps,
    SPFxCustomForm
}

/// <summary>
/// Represents a single list with its form customization type.
/// </summary>
public class CustomizedListItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public Guid ListId { get; set; }
    public string ListTitle { get; set; } = string.Empty;
    public string ListType { get; set; } = string.Empty;
    public ListFormType FormType { get; set; }
    public string FormTypeDescription => FormType switch
    {
        ListFormType.Default => "Default",
        ListFormType.PowerApps => "Power Apps",
        ListFormType.SPFxCustomForm => "SPFx Custom Form",
        _ => FormType.ToString()
    };
    public int ItemCount { get; set; }
    public string ListUrl { get; set; } = string.Empty;
    public string DefaultNewFormUrl { get; set; } = string.Empty;
    public string DefaultEditFormUrl { get; set; } = string.Empty;
    public string DefaultDisplayFormUrl { get; set; } = string.Empty;
    public string SpfxNewFormComponentId { get; set; } = string.Empty;
    public string SpfxEditFormComponentId { get; set; } = string.Empty;
    public string SpfxDisplayFormComponentId { get; set; } = string.Empty;
    public bool IsCustomized => FormType != ListFormType.Default;
}

/// <summary>
/// Represents customized list results for a single SharePoint site.
/// </summary>
public class SiteCustomizedListsResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<CustomizedListItem> Lists { get; set; } = [];
    public int TotalLists => Lists.Count;
    public int CustomizedCount => Lists.Count(l => l.IsCustomized);
    public int PowerAppsCount => Lists.Count(l => l.FormType == ListFormType.PowerApps);
    public int SpfxCount => Lists.Count(l => l.FormType == ListFormType.SPFxCustomForm);
}

/// <summary>
/// Represents the complete result of a customized lists report task execution.
/// </summary>
public class CustomizedListsReportResult
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
    public List<SiteCustomizedListsResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    /// <summary>
    /// Gets all lists flattened across all sites.
    /// </summary>
    public IEnumerable<CustomizedListItem> GetAllLists()
    {
        foreach (var siteResult in SiteResults)
        {
            foreach (var list in siteResult.Lists)
            {
                yield return list;
            }
        }
    }

    /// <summary>
    /// Gets only customized lists (Power Apps or SPFx) flattened across all sites.
    /// </summary>
    public IEnumerable<CustomizedListItem> GetCustomizedLists()
    {
        return GetAllLists().Where(l => l.IsCustomized);
    }

    public int TotalListsScanned => SiteResults.Sum(s => s.TotalLists);
    public int TotalCustomized => SiteResults.Sum(s => s.CustomizedCount);
    public int TotalPowerApps => SiteResults.Sum(s => s.PowerAppsCount);
    public int TotalSpfx => SiteResults.Sum(s => s.SpfxCount);

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }
}

/// <summary>
/// Export model for customized lists report CSV.
/// </summary>
public class CustomizedListsExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string ListTitle { get; set; } = string.Empty;
    public string ListType { get; set; } = string.Empty;
    public string FormType { get; set; } = string.Empty;
    public int ItemCount { get; set; }
    public string ListUrl { get; set; } = string.Empty;
    public string DefaultNewFormUrl { get; set; } = string.Empty;
    public string DefaultEditFormUrl { get; set; } = string.Empty;
    public string SpfxNewFormComponentId { get; set; } = string.Empty;
    public string SpfxEditFormComponentId { get; set; } = string.Empty;
    public string SpfxDisplayFormComponentId { get; set; } = string.Empty;
}

/// <summary>
/// Export model for customized lists site summary CSV.
/// </summary>
public class CustomizedListsSiteSummaryExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public int TotalLists { get; set; }
    public int CustomizedCount { get; set; }
    public int PowerAppsCount { get; set; }
    public int SpfxCount { get; set; }
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

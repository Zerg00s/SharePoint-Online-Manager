namespace SharePointOnlineManager.Models;

/// <summary>
/// Defines the type of task that can be executed.
/// </summary>
public enum TaskType
{
    /// <summary>
    /// Report on all lists across selected site collections.
    /// </summary>
    ListsReport,

    /// <summary>
    /// Compare list item counts between source and target sites.
    /// </summary>
    ListCompare,

    /// <summary>
    /// Report on all documents/files across selected site collections.
    /// </summary>
    DocumentReport,

    /// <summary>
    /// Report on permissions across selected site collections.
    /// </summary>
    PermissionReport,

    /// <summary>
    /// Set site state (Unlock, ReadOnly, NoAccess) for selected site collections.
    /// </summary>
    SetSiteState,

    /// <summary>
    /// Add site collection administrators to selected site collections.
    /// </summary>
    AddSiteCollectionAdmins,

    /// <summary>
    /// Remove site collection administrators from selected site collections.
    /// </summary>
    RemoveSiteCollectionAdmins,

    /// <summary>
    /// Compare and sync navigation settings (HorizontalQuickLaunch, MegaMenuEnabled) between source and target sites.
    /// </summary>
    NavigationSettingsSync,

    /// <summary>
    /// Compare documents between source and target sites to identify missing or mismatched files.
    /// </summary>
    DocumentCompare,

    /// <summary>
    /// Check site access for source and target accounts to identify permission issues.
    /// </summary>
    SiteAccessCheck,

    /// <summary>
    /// Report on external ad hoc (OTP) guest users across selected site collections.
    /// </summary>
    AdHocUsersReport,

    /// <summary>
    /// Find lists customized with Power Apps or SPFx forms across selected site collections.
    /// </summary>
    CustomizedListsReport,

    /// <summary>
    /// Find sites where SharePoint Publishing features are activated.
    /// </summary>
    PublishingSitesReport,

    /// <summary>
    /// Report on custom (non-OOTB) fields across lists and libraries.
    /// </summary>
    CustomFieldsReport,

    /// <summary>
    /// Report on site collections that have subsites.
    /// </summary>
    SubsitesReport,

    /// <summary>
    /// Find broken OneNote notebooks and optionally fix them.
    /// </summary>
    BrokenOneNoteReport
}

/// <summary>
/// Defines the execution status of a task.
/// </summary>
public enum TaskStatus
{
    Pending,
    Running,
    Completed,
    Failed,
    Cancelled
}

/// <summary>
/// Represents a persistent task definition that can be saved and re-executed.
/// </summary>
public class TaskDefinition
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = string.Empty;
    public TaskType Type { get; set; }
    public Guid ConnectionId { get; set; }
    public List<string> TargetSiteUrls { get; set; } = [];
    public string? ConfigurationJson { get; set; }
    public TaskStatus Status { get; set; } = TaskStatus.Pending;
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime? LastRunAt { get; set; }
    public DateTime? CompletedAt { get; set; }
    public string? LastError { get; set; }
    public int TotalSites => TargetSiteUrls.Count;

    /// <summary>
    /// Gets the task type as a human-readable string.
    /// </summary>
    public string TypeDescription => Type switch
    {
        TaskType.ListsReport => "Lists Report",
        TaskType.ListCompare => "List Compare",
        TaskType.DocumentReport => "Document Report",
        TaskType.PermissionReport => "Permission Report",
        TaskType.SetSiteState => "Set Site State",
        TaskType.AddSiteCollectionAdmins => "Add Site Admins",
        TaskType.RemoveSiteCollectionAdmins => "Remove Site Admins",
        TaskType.NavigationSettingsSync => "Navigation Settings",
        TaskType.DocumentCompare => "Document Compare",
        TaskType.SiteAccessCheck => "Site Access Check",
        TaskType.AdHocUsersReport => "Ad Hoc Users Report",
        TaskType.CustomizedListsReport => "Customized Lists Report",
        TaskType.PublishingSitesReport => "Publishing Sites Report",
        TaskType.CustomFieldsReport => "Custom Fields Report",
        TaskType.SubsitesReport => "Subsites Report",
        TaskType.BrokenOneNoteReport => "Broken OneNote Report",
        _ => Type.ToString()
    };

    /// <summary>
    /// Gets the status as a human-readable string.
    /// </summary>
    public string StatusDescription => Status switch
    {
        TaskStatus.Pending => "Pending",
        TaskStatus.Running => "Running",
        TaskStatus.Completed => "Completed",
        TaskStatus.Failed => "Failed",
        TaskStatus.Cancelled => "Cancelled",
        _ => Status.ToString()
    };
}

/// <summary>
/// Extension methods for TaskType enum.
/// </summary>
public static class TaskTypeExtensions
{
    /// <summary>
    /// Gets a human-readable description of the task type.
    /// </summary>
    public static string GetDescription(this TaskType type) => type switch
    {
        TaskType.ListsReport => "Report on all lists across selected sites",
        TaskType.ListCompare => "Compare list item counts between source and target sites",
        TaskType.DocumentReport => "Report on all documents/files across selected sites",
        TaskType.PermissionReport => "Report on permissions across selected sites",
        TaskType.SetSiteState => "Set site state: Unlock, ReadOnly, or NoAccess",
        TaskType.AddSiteCollectionAdmins => "Add up to 5 site collection administrators",
        TaskType.RemoveSiteCollectionAdmins => "Remove up to 5 site collection administrators",
        TaskType.NavigationSettingsSync => "Compare and sync navigation settings between tenants",
        TaskType.DocumentCompare => "Compare documents between source and target sites",
        TaskType.SiteAccessCheck => "Check site access for source and target accounts",
        TaskType.AdHocUsersReport => "Report on external ad hoc (OTP) guest users across sites",
        TaskType.CustomizedListsReport => "Find lists customized with Power Apps or SPFx forms",
        TaskType.PublishingSitesReport => "Find sites with SharePoint Publishing features activated",
        TaskType.CustomFieldsReport => "Report on custom (non-OOTB) fields across lists and libraries",
        TaskType.SubsitesReport => "Report on site collections that have one or more subsites",
        TaskType.BrokenOneNoteReport => "Find broken OneNote notebooks and optionally fix them",
        _ => type.ToString()
    };

    /// <summary>
    /// Gets the display name for the task type.
    /// </summary>
    public static string GetDisplayName(this TaskType type) => type switch
    {
        TaskType.ListsReport => "Lists Report",
        TaskType.ListCompare => "List Compare",
        TaskType.DocumentReport => "Document Report",
        TaskType.PermissionReport => "Permission Report",
        TaskType.SetSiteState => "Set Site State",
        TaskType.AddSiteCollectionAdmins => "Add Site Collection Administrators",
        TaskType.RemoveSiteCollectionAdmins => "Remove Site Collection Administrators",
        TaskType.NavigationSettingsSync => "Navigation Settings Sync",
        TaskType.DocumentCompare => "Document Compare",
        TaskType.SiteAccessCheck => "Site Access Check",
        TaskType.AdHocUsersReport => "Ad Hoc Users Report",
        TaskType.CustomizedListsReport => "Customized Lists Report",
        TaskType.PublishingSitesReport => "Publishing Sites Report",
        TaskType.CustomFieldsReport => "Custom Fields Report",
        TaskType.SubsitesReport => "Subsites Report",
        TaskType.BrokenOneNoteReport => "Broken OneNote Report",
        _ => type.ToString()
    };

    /// <summary>
    /// Indicates whether the task type requires dual connections (source and target).
    /// </summary>
    public static bool RequiresDualConnections(this TaskType type) => type switch
    {
        TaskType.ListCompare => true,
        TaskType.NavigationSettingsSync => true,
        TaskType.DocumentCompare => true,
        TaskType.SiteAccessCheck => true,
        _ => false
    };
}

namespace SharePointOnlineManager.Models;

/// <summary>
/// Well-known out-of-the-box SharePoint field groups.
/// Fields in these groups are NOT considered custom.
/// </summary>
public static class OotbFieldGroups
{
    public static readonly HashSet<string> Groups = new(StringComparer.OrdinalIgnoreCase)
    {
        "_Hidden",
        "Base Columns",
        "Content Feedback",
        "Core Contact and Calendar Columns",
        "Core Document Columns",
        "Core Task and Issue Columns",
        "Display Template Columns",
        "Document and Record Management Columns",
        "Enterprise Keywords Group",
        "Extended Columns",
        "JavaScript Display Template",
        "Page Layout Columns",
        "Publishing Columns",
        "Ratings",
        "Reports",
        "Status Indicators"
    };

    /// <summary>
    /// Well-known system field internal names that SharePoint places in the "Custom Columns" group
    /// despite being built-in. These are excluded from custom field reports.
    /// </summary>
    public static readonly HashSet<string> SystemFieldInternalNames = new(StringComparer.OrdinalIgnoreCase)
    {
        // Core system fields
        "ID", "Title", "Created", "Modified", "Author", "Editor",
        "FileLeafRef", "FileDirRef", "FileRef", "File_x0020_Type",
        "_CopySource", "_CheckinComment", "_UIVersionString",

        // Computed/display fields
        "LinkFilenameNoMenu", "LinkFilename", "LinkFilename2", "DocIcon",
        "FileSizeDisplay", "Edit", "SelectTitle",

        // Child count / folder fields
        "ItemChildCount", "FolderChildCount",

        // Compliance / retention / sensitivity labels
        "_ComplianceFlags", "_ComplianceTag", "_ComplianceTagWrittenTime",
        "_ComplianceTagUserId", "_IsRecord", "_DisplayName",
        "ComplianceAssetId",

        // Social fields
        "_CommentCount", "_LikeCount", "_CommentFlags",

        // App fields
        "AppAuthor", "AppEditor",

        // Lookup / versioning
        "ParentVersionString", "ParentLeafName",
        "CheckoutUser", "CheckedOutTitle",

        // Document ID fields (OOTB Document ID Service feature)
        "_dlc_DocId", "_dlc_DocIdUrl", "_dlc_DocIdPersistId",

        // Color tag
        "_ColorTag", "_ColorHex",

        // Other system fields commonly in "Custom Columns"
        "ContentType", "ContentTypeId",
        "SMTotalSize", "SMLastModifiedDate", "SMTotalFileStreamSize", "SMTotalFileCount",
        "ProgId", "ScopeId", "VirusStatus",
        "_ShortcutUrl", "_ShortcutSiteId", "_ShortcutWebId", "_ShortcutUniqueId",
        "AccessPolicy", "BSN",
        "SyncClientId", "TemplateUrl",
        "A2ODMountCount",
        "Restricted", "OriginatorId",
        "NoExecute", "ContentVersion",
        "ComplianceFlags",
        "MediaServiceMetadata", "MediaServiceFastMetadata",
        "MediaServiceAutoTags", "MediaServiceOCR",
        "MediaServiceGenerationTime", "MediaServiceEventHashCode",
        "MediaServiceDateTaken", "MediaServiceAutoKeyPoints",
        "MediaServiceKeyPoints", "MediaServiceLocation",
        "MediaLengthInSeconds",
        "MediaServiceSearchProperties",
        "LCF", "TaxCatchAll", "TaxCatchAllLabel",
        "_ExtendedDescription",

        // ── Events / Calendar list template fields ──
        "Location", "Geolocation",
        "EventDate", "EndDate", "Duration",
        "fAllDayEvent", "fRecurrence",
        "RecurrenceID", "RecurrenceData",
        "TimeZone", "XMLTZone",
        "MasterSeriesItemID", "Workspace", "WorkspaceLink",
        "EventType", "UID", "EventCanceled",
        "ParticipantsPicker", "Category", "Facilities",
        "FreeBusy", "Overbook",
        "BannerUrl",
        "Description",

        // ── Site Pages / Wiki library fields ──
        "WikiField", "CanvasContent1", "BannerImageUrl",
        "PromotedState", "FirstPublishedDate",
        "LayoutWebpartsContent", "_OriginalSourceUrl",
        "_OriginalSourceSiteId", "_OriginalSourceWebId",
        "_OriginalSourceListId", "_OriginalSourceItemId",
        "_AuthorBylineId", "PageLayoutType",
        "_TopicHeader", "_SPSitePageFlags",
        "_SPCallToAction",

        // ── Tasks list template fields ──
        "StartDate", "DueDate",
        "PercentComplete", "Priority",
        "TaskStatus", "AssignedTo",
        "Body", "Predecessors",
        "TaskGroup", "TaskOutcome",
        "RelatedItems",

        // ── Contacts list template fields ──
        "FirstName", "FullName",
        "Company", "JobTitle",
        "WorkPhone", "HomePhone", "CellPhone",
        "WorkFax", "WorkAddress", "WorkCity",
        "WorkState", "WorkZip", "WorkCountry",
        "WebPage", "Comments",

        // ── Announcements list template fields ──
        "Expires",

        // ── Discussion Board fields ──
        "DiscussionTitle", "DiscussionLastUpdated",
        "Threading", "ThreadIndex", "ThreadTopic",
        "ParentFolderId", "BodyAndMore", "CorrectBodyToShow",
        "TrimmedBody", "IsRootPost", "ReplyCount",
        "ParentItemID", "ParentItemEditor", "LastReplyBy",
        "IsQuestion", "BestAnswerId", "IsFeatured",

        // ── Survey fields ──
        "Completed",

        // ── Issue Tracking fields ──
        "IssueStatus", "V3Comments", "RelatedIssues",

        // ── Document library system fields ──
        "SharedWithUsers", "SharedWithDetails",
        "SharingHintHash", "_VirusInfo", "_VirusVendorID", "_VirusStatus",
        "CheckInComment",

        // ── Common OOTB fields across multiple templates ──
        "Attachments", "Order", "GUID",
        "WorkflowVersion", "WorkflowInstanceID",
        "SelectFilename", "InstanceID",
        "UniqueId", "SyncClientId", "MetaInfo",
        "owshiddenversion", "_Level", "_IsCurrentVersion",
        "Last_x0020_Modified", "Created_x0020_Date",
        "FSObjType", "SortBehavior", "PermMask",
        "PrincipalCount", "FileLeafRef", "FileSizeDisplay",
        "ParentUniqueId", "StreamHash",
        "_HasCopyDestinations", "_CopySource",
        "_ModerationStatus", "_ModerationComments"
    };

    /// <summary>
    /// Returns true if the field group indicates a custom (non-OOTB) field.
    /// Custom = "Custom Columns" group OR any group not in the OOTB set.
    /// </summary>
    public static bool IsCustomGroup(string? group)
    {
        if (string.IsNullOrEmpty(group))
            return false;

        // "Custom Columns" is the default group for user-created site columns
        if (group.Equals("Custom Columns", StringComparison.OrdinalIgnoreCase))
            return true;

        // Any group not in the OOTB set is custom
        return !Groups.Contains(group);
    }

    /// <summary>
    /// Returns true if the internal name matches a known system/built-in field
    /// that SharePoint places in the "Custom Columns" group.
    /// </summary>
    public static bool IsSystemFieldInternalName(string internalName)
    {
        if (string.IsNullOrEmpty(internalName))
            return false;

        return SystemFieldInternalNames.Contains(internalName);
    }
}

/// <summary>
/// Represents a single custom field found on a list.
/// </summary>
public class CustomFieldItem
{
    public string SiteCollectionUrl { get; set; } = string.Empty;
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string ListTitle { get; set; } = string.Empty;
    public string ListUrl { get; set; } = string.Empty;
    public int ItemCount { get; set; }
    public string ColumnName { get; set; } = string.Empty;
    public string InternalName { get; set; } = string.Empty;
    public string FieldType { get; set; } = string.Empty;
    public string Group { get; set; } = string.Empty;
    public DateTime ListCreated { get; set; }
    public DateTime ListModified { get; set; }
}

/// <summary>
/// Represents the custom fields result for a single site.
/// </summary>
public class SiteCustomFieldsResult
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public bool Success { get; set; }
    public string? ErrorMessage { get; set; }
    public List<CustomFieldItem> Fields { get; set; } = [];
    public int ListsScanned { get; set; }
    public int ListsWithCustomFields => Fields.Select(f => f.ListTitle).Distinct(StringComparer.OrdinalIgnoreCase).Count();
    public int CustomFieldCount => Fields.Count;
}

/// <summary>
/// Represents the complete result of a custom fields report task execution.
/// </summary>
public class CustomFieldsReportResult
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
    public List<SiteCustomFieldsResult> SiteResults { get; set; } = [];
    public List<string> ExecutionLog { get; set; } = [];

    public int TotalCustomFields => SiteResults.Where(s => s.Success).Sum(s => s.CustomFieldCount);
    public int TotalListsWithCustomFields => SiteResults.Where(s => s.Success).Sum(s => s.ListsWithCustomFields);
    public int TotalListsScanned => SiteResults.Where(s => s.Success).Sum(s => s.ListsScanned);

    /// <summary>
    /// Returns all custom fields flattened across all sites.
    /// </summary>
    public IEnumerable<CustomFieldItem> GetAllFields() =>
        SiteResults.Where(s => s.Success).SelectMany(s => s.Fields);

    /// <summary>
    /// Returns distinct field groups found across all custom fields.
    /// </summary>
    public IEnumerable<string> GetDistinctGroups() =>
        GetAllFields().Select(f => f.Group).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(g => g);

    /// <summary>
    /// Adds a log entry with timestamp.
    /// </summary>
    public void Log(string message)
    {
        ExecutionLog.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
    }
}

/// <summary>
/// Export model for custom fields report CSV.
/// </summary>
public class CustomFieldsExportItem
{
    public string SiteCollectionUrl { get; set; } = string.Empty;
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public string ListTitle { get; set; } = string.Empty;
    public string ListUrl { get; set; } = string.Empty;
    public int ItemCount { get; set; }
    public string ColumnName { get; set; } = string.Empty;
    public string InternalName { get; set; } = string.Empty;
    public string FieldType { get; set; } = string.Empty;
    public string Group { get; set; } = string.Empty;
    public DateTime ListCreated { get; set; }
    public DateTime ListModified { get; set; }
}

/// <summary>
/// Export model for custom fields site summary CSV.
/// </summary>
public class CustomFieldsSiteSummaryExportItem
{
    public string SiteUrl { get; set; } = string.Empty;
    public string SiteTitle { get; set; } = string.Empty;
    public int ListsScanned { get; set; }
    public int ListsWithCustomFields { get; set; }
    public int CustomFieldCount { get; set; }
    public string Status { get; set; } = string.Empty;
    public string Error { get; set; } = string.Empty;
}

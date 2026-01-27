namespace SharePointOnlineManager.Models;

/// <summary>
/// Defines the type of SharePoint site.
/// </summary>
public enum SiteType
{
    Unknown,
    TeamSite,
    CommunicationSite,
    OneDrive
}

/// <summary>
/// Represents a site collection retrieved from the SharePoint Admin API.
/// </summary>
public class SiteCollection
{
    public string Url { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Template { get; set; } = string.Empty;
    public string Owner { get; set; } = string.Empty;
    public long StorageUsed { get; set; }
    public long StorageQuota { get; set; }
    public DateTime CreatedDate { get; set; }
    public DateTime LastModifiedDate { get; set; }
    public int WebsCount { get; set; }
    public string Status { get; set; } = string.Empty;
    public DateTime LastActivityDate { get; set; }
    public int FileCount { get; set; }
    public string ExternalSharing { get; set; } = string.Empty;
    public Guid HubSiteId { get; set; }
    public Guid GroupId { get; set; }
    public int PageViews { get; set; }
    public int PagesVisited { get; set; }
    public int State { get; set; }
    public DateTime? TimeDeleted { get; set; }
    public int ChannelType { get; set; }
    public int LanguageLcid { get; set; }

    /// <summary>
    /// Gets the language display name from the LCID.
    /// </summary>
    public string LanguageDisplay => GetLanguageDisplayName(LanguageLcid);

    /// <summary>
    /// Gets a formatted last activity display string.
    /// </summary>
    public string LastActivityDisplay => LastActivityDate == DateTime.MinValue
        ? string.Empty
        : LastActivityDate.ToString("yyyy-MM-dd");

    /// <summary>
    /// Gets whether the site is connected to a Microsoft 365 group.
    /// </summary>
    public bool IsGroupConnected => GroupId != Guid.Empty;

    /// <summary>
    /// Gets whether the site is associated with a hub site.
    /// </summary>
    public bool IsHubAssociated => HubSiteId != Guid.Empty;

    /// <summary>
    /// Gets the hub status display string.
    /// </summary>
    public string HubDisplay
    {
        get
        {
            if (HubSiteId == Guid.Empty)
                return "";
            // If the site's own ID matches the HubSiteId, it IS the hub
            if (SiteId == HubSiteId)
                return "Hub";
            return "Yes";
        }
    }

    /// <summary>
    /// The site's own unique identifier.
    /// </summary>
    public Guid SiteId { get; set; }

    /// <summary>
    /// Gets a human-readable state description.
    /// </summary>
    public string StateDisplay => State switch
    {
        0 => "Active",
        1 => "Active",
        2 => "Locked",
        3 => "NoAccess",
        _ => State.ToString()
    };

    /// <summary>
    /// Gets whether the site is deleted (in recycle bin).
    /// </summary>
    public bool IsDeleted => TimeDeleted.HasValue;

    /// <summary>
    /// Gets a description of the channel type.
    /// </summary>
    public string ChannelTypeDisplay => ChannelType switch
    {
        0 => "",
        1 => "Private",
        2 => "Shared",
        _ => $"Channel({ChannelType})"
    };

    /// <summary>
    /// Gets the connection type of the site (Standalone, Group, Teams Channel).
    /// </summary>
    public string ConnectionType
    {
        get
        {
            if (ChannelType > 0)
                return ChannelType == 1 ? "Private Channel" : "Shared Channel";

            if (GroupId != Guid.Empty)
                return "Group/Teams";

            return "Standalone";
        }
    }

    /// <summary>
    /// Gets the site type based on template and URL patterns.
    /// </summary>
    public SiteType SiteType
    {
        get
        {
            if (Template.Contains("SPSPERS", StringComparison.OrdinalIgnoreCase) ||
                Url.Contains("-my.sharepoint.com/personal", StringComparison.OrdinalIgnoreCase))
            {
                return SiteType.OneDrive;
            }

            if (Template.Equals("SITEPAGEPUBLISHING#0", StringComparison.OrdinalIgnoreCase))
            {
                return SiteType.CommunicationSite;
            }

            if (Template.Contains("GROUP", StringComparison.OrdinalIgnoreCase) ||
                Template.Equals("STS#3", StringComparison.OrdinalIgnoreCase) ||
                Template.Equals("STS#0", StringComparison.OrdinalIgnoreCase))
            {
                return SiteType.TeamSite;
            }

            return SiteType.Unknown;
        }
    }

    /// <summary>
    /// Gets a human-readable site type description.
    /// </summary>
    public string SiteTypeDescription => SiteType switch
    {
        SiteType.OneDrive => "OneDrive",
        SiteType.CommunicationSite => "Communication Site",
        SiteType.TeamSite => "Team Site",
        _ => "SharePoint Site"
    };

    /// <summary>
    /// Gets storage used formatted as a human-readable string.
    /// </summary>
    public string StorageUsedFormatted => FormatBytes(StorageUsed);

    /// <summary>
    /// Gets storage used in bytes (alias for StorageUsed for clarity in UI code).
    /// </summary>
    public long StorageUsedBytes => StorageUsed;

    private static string FormatBytes(long bytes)
    {
        string[] suffixes = ["B", "KB", "MB", "GB", "TB"];
        int suffixIndex = 0;
        double size = bytes;

        while (size >= 1024 && suffixIndex < suffixes.Length - 1)
        {
            size /= 1024;
            suffixIndex++;
        }

        return $"{size:N2} {suffixes[suffixIndex]}";
    }

    private static string GetLanguageDisplayName(int lcid)
    {
        if (lcid == 0)
            return string.Empty;

        // Common SharePoint language LCIDs
        return lcid switch
        {
            1025 => "Arabic",
            1026 => "Bulgarian",
            1027 => "Catalan",
            1028 => "Chinese (Traditional)",
            1029 => "Czech",
            1030 => "Danish",
            1031 => "German",
            1032 => "Greek",
            1033 => "English",
            1035 => "Finnish",
            1036 => "French",
            1037 => "Hebrew",
            1038 => "Hungarian",
            1040 => "Italian",
            1041 => "Japanese",
            1042 => "Korean",
            1043 => "Dutch",
            1044 => "Norwegian (Bokmal)",
            1045 => "Polish",
            1046 => "Portuguese (Brazil)",
            1048 => "Romanian",
            1049 => "Russian",
            1050 => "Croatian",
            1051 => "Slovak",
            1053 => "Swedish",
            1054 => "Thai",
            1055 => "Turkish",
            1057 => "Indonesian",
            1058 => "Ukrainian",
            1060 => "Slovenian",
            1061 => "Estonian",
            1062 => "Latvian",
            1063 => "Lithuanian",
            1066 => "Vietnamese",
            1081 => "Hindi",
            1086 => "Malay",
            2052 => "Chinese (Simplified)",
            2070 => "Portuguese (Portugal)",
            2074 => "Serbian (Latin)",
            3082 => "Spanish",
            _ => lcid.ToString()
        };
    }

    /// <summary>
    /// Gets a human-readable description of the site template.
    /// </summary>
    public string TemplateDescription => GetTemplateDescription(Template);

    private static string GetTemplateDescription(string template)
    {
        if (string.IsNullOrEmpty(template))
            return string.Empty;

        return template.ToUpperInvariant() switch
        {
            "ACCSRV#0" => "Access Services Site",
            "ACCSVC#0" => "Access Services Site Internal",
            "ACCSVC#1" => "Access Services Site",
            "APP#0" => "App Template",
            "APPCATALOG#0" => "App Catalog Site",
            "BDR#0" => "Document Center",
            "BICENTERSITE#0" => "Business Intelligence Center",
            "BLANKINTERNET#0" => "Publishing Site",
            "BLANKINTERNET#1" => "Press Releases Site",
            "BLANKINTERNET#2" => "Publishing Site with Workflow",
            "BLANKINTERNETCONTAINER#0" => "Publishing Portal",
            "BLOG#0" => "Blog",
            "CENTRALADMIN#0" => "Central Admin Site",
            "CMSPUBLISHING#0" => "Publishing Site",
            "COMMUNITY#0" => "Community Site",
            "COMMUNITYPORTAL#0" => "Community Portal",
            "DEV#0" => "Developer Site",
            "EHS#1" => "Team Site - SharePoint Online (Legacy)",
            "EDISC#0" => "eDiscovery Center",
            "EDISC#1" => "eDiscovery Case",
            "ENTERWIKI#0" => "Enterprise Wiki",
            "GLOBAL#0" => "Global Template",
            "GROUP#0" => "Team Site (Microsoft 365 Group)",
            "MPS#0" => "Basic Meeting Workspace",
            "MPS#1" => "Blank Meeting Workspace",
            "MPS#2" => "Decision Meeting Workspace",
            "MPS#3" => "Social Meeting Workspace",
            "MPS#4" => "Multipage Meeting Workspace",
            "OFFILE#0" => "Records Center (Obsolete)",
            "OFFILE#1" => "Records Center",
            "OSRV#0" => "Shared Services Administration Site",
            "POINTPUBLISHINGHUB#0" => "Point Publishing Hub",
            "POINTPUBLISHINGPERSONAL#0" => "Point Publishing Personal",
            "POINTPUBLISHINGTOPIC#0" => "Video Channel",
            "POLICYCTR#0" => "Compliance Policy Center",
            "PPSMASITE#0" => "PerformancePoint",
            "PRODUCTCATALOG#0" => "Product Catalog",
            "PROFILES#0" => "Profiles",
            "PROJECTSITE#0" => "Project Site",
            "PWA#0" => "Project Web App Site",
            "PWS#0" => "Microsoft Project Site",
            "REDIRECTSITE#0" => "Redirect Site",
            "REDIRECTSITE#1" => "Redirect Site",
            "REVIEWCTR#0" => "Review Center for Retention",
            "REVIEWCTR#1" => "Review Center for Retention",
            "SGS#0" => "Group Work Site",
            "SPS#0" => "SharePoint Portal Server Site",
            "SPSCOMMU#0" => "Community Area Template",
            "SPSMSITE#0" => "Personalization Site",
            "SPSMSITEHOST#0" => "My Site Host",
            "SPSNEWS#0" => "News Site",
            "SPSNHOME#0" => "News Site",
            "SPSPERS#0" => "OneDrive for Business",
            "SPSPERS#2" => "OneDrive for Business",
            "SPSPERS#3" => "OneDrive for Business (Storage Only)",
            "SPSPERS#4" => "OneDrive for Business (Social Only)",
            "SPSPERS#5" => "OneDrive for Business (Empty)",
            "SPSPERS#6" => "OneDrive for Business",
            "SPSPERS#7" => "OneDrive for Business",
            "SPSPERS#8" => "OneDrive for Business",
            "SPSPERS#9" => "OneDrive for Business",
            "SPSPERS#10" => "OneDrive for Business",
            "SPSPORTAL#0" => "Collaboration Portal",
            "SPSREPORTCENTER#0" => "Report Center",
            "SPSSITES#0" => "Site Directory",
            "SPSTOC#0" => "Contents Area Template",
            "SPSTOPIC#0" => "Topic Area Template",
            "SRCHCEN#0" => "Enterprise Search Center",
            "SRCHCENTERLITE#0" => "Basic Search Center",
            "SRCHCENTERLITE#1" => "Basic Search Center",
            "SITEPAGEPUBLISHING#0" => "Communication Site",
            "STS#0" => "Classic Team Site",
            "STS#1" => "Blank Site",
            "STS#2" => "Document Workspace",
            "STS#3" => "Modern Team Site",
            "TBH#0" => "In-Place Hold Policy Center",
            "TENANTADMIN#0" => "Tenant Admin Site",
            "TEAMCHANNEL#0" => "Teams Channel Site",
            "TEAMCHANNEL#1" => "Teams Channel Site",
            "VISPRUS#0" => "Visio Process Repository",
            "WIKI#0" => "Wiki Site",
            _ => template // Return the template code if not found
        };
    }
}

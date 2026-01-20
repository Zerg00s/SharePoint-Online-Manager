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
}

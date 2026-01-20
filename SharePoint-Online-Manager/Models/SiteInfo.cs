namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents information about a SharePoint site.
/// </summary>
public class SiteInfo
{
    public string Url { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string WebTemplate { get; set; } = string.Empty;
    public DateTime Created { get; set; }
    public DateTime LastItemModifiedDate { get; set; }
    public bool IsConnected { get; set; }
    public string? ErrorMessage { get; set; }
}

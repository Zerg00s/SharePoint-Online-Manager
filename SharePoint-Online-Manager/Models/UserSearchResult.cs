namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents a user search result from SharePoint people picker.
/// </summary>
public class UserSearchResult
{
    /// <summary>
    /// The display name of the user.
    /// </summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// The email address of the user.
    /// </summary>
    public string Email { get; set; } = string.Empty;

    /// <summary>
    /// The login name for API calls (e.g., "i:0#.f|membership|user@domain.com").
    /// </summary>
    public string LoginName { get; set; } = string.Empty;

    /// <summary>
    /// The entity type (User, Group, etc.).
    /// </summary>
    public string EntityType { get; set; } = string.Empty;

    public override string ToString() => !string.IsNullOrEmpty(Email)
        ? $"{DisplayName} ({Email})"
        : DisplayName;
}

/// <summary>
/// Configuration for Add Site Collection Administrators task.
/// </summary>
public class AddSiteAdminsConfiguration
{
    /// <summary>
    /// The list of administrators to add.
    /// </summary>
    public List<UserSearchResult> Administrators { get; set; } = [];
}

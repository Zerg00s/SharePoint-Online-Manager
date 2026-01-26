namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents the authentication cookies captured from SharePoint Online login.
/// </summary>
public class AuthCookies
{
    public string Domain { get; set; } = string.Empty;
    public string FedAuth { get; set; } = string.Empty;
    public string RtFa { get; set; } = string.Empty;
    public string UserEmail { get; set; } = string.Empty;
    public DateTime CapturedAt { get; set; } = DateTime.UtcNow;

    public bool IsValid => !string.IsNullOrEmpty(FedAuth) && !string.IsNullOrEmpty(RtFa);
}

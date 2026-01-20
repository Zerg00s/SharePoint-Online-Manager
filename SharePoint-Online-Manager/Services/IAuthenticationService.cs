using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Interface for SharePoint authentication operations with multi-tenant support.
/// </summary>
public interface IAuthenticationService
{
    /// <summary>
    /// Gets the stored authentication cookies for the specified domain.
    /// </summary>
    AuthCookies? GetStoredCookies(string domain);

    /// <summary>
    /// Stores authentication cookies for a domain.
    /// </summary>
    void StoreCookies(AuthCookies cookies);

    /// <summary>
    /// Clears stored credentials for a specific domain.
    /// </summary>
    void ClearCredentials(string domain);

    /// <summary>
    /// Clears all stored credentials for all domains.
    /// </summary>
    void ClearAllCredentials();

    /// <summary>
    /// Checks if credentials are stored for a specific domain.
    /// </summary>
    bool HasStoredCredentials(string domain);

    /// <summary>
    /// Checks if any credentials are stored.
    /// </summary>
    bool HasAnyStoredCredentials();

    /// <summary>
    /// Gets all domains that have stored credentials.
    /// </summary>
    List<string> GetStoredDomains();
}

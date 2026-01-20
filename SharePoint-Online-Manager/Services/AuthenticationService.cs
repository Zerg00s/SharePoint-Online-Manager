using SharePointOnlineManager.Authentication;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Manages SharePoint authentication using cookie-based approach with multi-tenant support.
/// </summary>
public class AuthenticationService : IAuthenticationService
{
    private readonly CookieStore _cookieStore;

    public AuthenticationService()
    {
        _cookieStore = new CookieStore();
    }

    public AuthenticationService(CookieStore cookieStore)
    {
        _cookieStore = cookieStore;
    }

    public AuthCookies? GetStoredCookies(string domain)
    {
        return _cookieStore.Load(domain);
    }

    public void StoreCookies(AuthCookies cookies)
    {
        _cookieStore.Save(cookies);
    }

    public void ClearCredentials(string domain)
    {
        _cookieStore.Clear(domain);
    }

    public void ClearAllCredentials()
    {
        _cookieStore.ClearAll();
    }

    public bool HasStoredCredentials(string domain)
    {
        return _cookieStore.HasStoredCookies(domain);
    }

    public bool HasAnyStoredCredentials()
    {
        return _cookieStore.HasAnyStoredCookies();
    }

    public List<string> GetStoredDomains()
    {
        return _cookieStore.GetStoredDomains();
    }
}

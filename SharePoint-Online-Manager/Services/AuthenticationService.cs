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
        System.Diagnostics.Debug.WriteLine($"[SPOManager] AuthService.GetStoredCookies called for domain: '{domain}'");

        // Try exact domain first
        var cookies = _cookieStore.Load(domain);

        // If not found and this is a SharePoint tenant domain, try the admin domain
        // (cookies from admin site work for regular tenant sites)
        if (cookies == null && domain.EndsWith(".sharepoint.com") && !domain.Contains("-admin"))
        {
            var adminDomain = domain.Replace(".sharepoint.com", "-admin.sharepoint.com");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] AuthService.GetStoredCookies - Trying admin domain fallback: '{adminDomain}'");
            cookies = _cookieStore.Load(adminDomain);
        }

        // Also try the reverse - if looking for admin domain, try tenant domain
        if (cookies == null && domain.Contains("-admin.sharepoint.com"))
        {
            var tenantDomain = domain.Replace("-admin.sharepoint.com", ".sharepoint.com");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] AuthService.GetStoredCookies - Trying tenant domain fallback: '{tenantDomain}'");
            cookies = _cookieStore.Load(tenantDomain);
        }

        System.Diagnostics.Debug.WriteLine($"[SPOManager] AuthService.GetStoredCookies result: {(cookies == null ? "null" : $"Domain={cookies.Domain}, Valid={cookies.IsValid}, User={cookies.UserEmail}")}");
        return cookies;
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

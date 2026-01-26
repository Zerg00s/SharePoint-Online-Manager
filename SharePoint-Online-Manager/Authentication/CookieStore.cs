using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Authentication;

/// <summary>
/// Provides persistent storage for SharePoint authentication cookies using Windows DPAPI encryption.
/// Supports multi-tenant scenarios with per-domain cookie files.
/// </summary>
public partial class CookieStore
{
    private static readonly string AppDataFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SharePointOnlineManager");

    private static readonly string CookiesFolder = Path.Combine(AppDataFolder, "cookies");

    [GeneratedRegex(@"[^a-zA-Z0-9\-\.]")]
    private static partial Regex InvalidFileCharsRegex();

    /// <summary>
    /// Saves authentication cookies encrypted with DPAPI for the specified domain.
    /// </summary>
    public void Save(AuthCookies cookies)
    {
        if (string.IsNullOrEmpty(cookies.Domain))
        {
            throw new ArgumentException("Domain is required for saving cookies.", nameof(cookies));
        }

        try
        {
            var json = JsonSerializer.Serialize(cookies);
            var data = System.Text.Encoding.UTF8.GetBytes(json);
            var encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);

            Directory.CreateDirectory(CookiesFolder);
            var filePath = GetCookieFilePath(cookies.Domain);
            File.WriteAllBytes(filePath, encrypted);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to save cookies: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Loads and decrypts stored authentication cookies for the specified domain.
    /// </summary>
    public AuthCookies? Load(string domain)
    {
        var filePath = GetCookieFilePath(domain);
        System.Diagnostics.Debug.WriteLine($"[SPOManager] CookieStore.Load - Looking for: '{domain}' at path: '{filePath}'");
        System.Diagnostics.Debug.WriteLine($"[SPOManager] CookieStore.Load - File exists: {File.Exists(filePath)}");

        if (!File.Exists(filePath))
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] CookieStore.Load - No cookie file found for '{domain}'");
            return null;
        }

        try
        {
            var encrypted = File.ReadAllBytes(filePath);
            var data = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
            var json = System.Text.Encoding.UTF8.GetString(data);
            var cookies = JsonSerializer.Deserialize<AuthCookies>(json);

            System.Diagnostics.Debug.WriteLine($"[SPOManager] CookieStore.Load - Loaded cookies, Domain in file: '{cookies?.Domain}'");

            if (cookies != null && cookies.Domain.Equals(domain, StringComparison.OrdinalIgnoreCase))
            {
                System.Diagnostics.Debug.WriteLine($"[SPOManager] CookieStore.Load - Domain matches, returning cookies");
                return cookies;
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] CookieStore.Load - Domain mismatch: requested '{domain}' vs stored '{cookies?.Domain}'");
            return null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] CookieStore.Load - Exception: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Checks if stored cookies exist for the specified domain.
    /// </summary>
    public bool HasStoredCookies(string domain)
    {
        var filePath = GetCookieFilePath(domain);
        return File.Exists(filePath);
    }

    /// <summary>
    /// Checks if any stored cookies exist.
    /// </summary>
    public bool HasAnyStoredCookies()
    {
        if (!Directory.Exists(CookiesFolder))
            return false;

        return Directory.GetFiles(CookiesFolder, "*.dat").Length > 0;
    }

    /// <summary>
    /// Deletes stored cookies for the specified domain.
    /// </summary>
    public void Clear(string domain)
    {
        var filePath = GetCookieFilePath(domain);
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }

    /// <summary>
    /// Deletes all stored cookies for all domains.
    /// </summary>
    public void ClearAll()
    {
        if (Directory.Exists(CookiesFolder))
        {
            foreach (var file in Directory.GetFiles(CookiesFolder, "*.dat"))
            {
                try
                {
                    File.Delete(file);
                }
                catch
                {
                    // Ignore individual file deletion errors
                }
            }
        }
    }

    /// <summary>
    /// Gets all domains that have stored cookies.
    /// </summary>
    public List<string> GetStoredDomains()
    {
        var domains = new List<string>();

        if (!Directory.Exists(CookiesFolder))
            return domains;

        foreach (var file in Directory.GetFiles(CookiesFolder, "*.dat"))
        {
            try
            {
                var encrypted = File.ReadAllBytes(file);
                var data = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                var json = System.Text.Encoding.UTF8.GetString(data);
                var cookies = JsonSerializer.Deserialize<AuthCookies>(json);

                if (cookies != null && !string.IsNullOrEmpty(cookies.Domain))
                {
                    domains.Add(cookies.Domain);
                }
            }
            catch
            {
                // Skip invalid files
            }
        }

        return domains;
    }

    private static string GetCookieFilePath(string domain)
    {
        var sanitizedDomain = InvalidFileCharsRegex().Replace(domain.ToLowerInvariant(), "_");
        return Path.Combine(CookiesFolder, $"{sanitizedDomain}.dat");
    }
}

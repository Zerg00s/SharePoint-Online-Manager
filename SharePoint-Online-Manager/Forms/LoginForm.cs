using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Forms;

/// <summary>
/// Login form that uses WebView2 to capture SharePoint authentication cookies.
/// </summary>
public partial class LoginForm : Form
{
    private WebView2 _webView = null!;
    private readonly string _siteUrl;
    private readonly string _domain;
    private bool _loginComplete;

    public AuthCookies? CapturedCookies { get; private set; }

    public LoginForm(string siteUrl)
    {
        _siteUrl = siteUrl;
        var uri = new Uri(siteUrl);
        _domain = uri.Host;

        InitializeComponent();
        InitializeWebView();
    }

    private void InitializeWebView()
    {
        _webView = new WebView2
        {
            Dock = DockStyle.Fill
        };
        Controls.Add(_webView);
    }

    protected override async void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        await InitializeAndNavigateAsync();
    }

    private async Task InitializeAndNavigateAsync()
    {
        try
        {
            // Initialize WebView2 with a custom user data folder
            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SharePointOnlineManager", "WebView2");

            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await _webView.EnsureCoreWebView2Async(env);

            // Set up navigation event handlers
            _webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;

            // Navigate to the SharePoint site to trigger login
            _webView.CoreWebView2.Navigate(_siteUrl);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to initialize browser: {ex.Message}\n\nMake sure WebView2 Runtime is installed.",
                "Browser Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }

    private async void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (_loginComplete)
            return;

        try
        {
            var currentUrl = _webView.Source?.ToString() ?? string.Empty;
            System.Diagnostics.Debug.WriteLine($"[SPOManager] Navigation completed: {currentUrl}");

            // Check if we've been redirected back to the SharePoint site (login complete)
            if (currentUrl.StartsWith(_siteUrl, StringComparison.OrdinalIgnoreCase) ||
                (currentUrl.Contains(_domain) && !currentUrl.Contains("login.microsoftonline.com")))
            {
                System.Diagnostics.Debug.WriteLine($"[SPOManager] Detected SharePoint site, capturing cookies for domain: {_domain}");

                // Try to capture cookies
                var cookies = await _webView.CoreWebView2.CookieManager.GetCookiesAsync($"https://{_domain}");
                System.Diagnostics.Debug.WriteLine($"[SPOManager] Found {cookies.Count} cookies");

                var fedAuth = cookies.FirstOrDefault(c => c.Name == "FedAuth")?.Value;
                var rtFa = cookies.FirstOrDefault(c => c.Name == "rtFa")?.Value;

                System.Diagnostics.Debug.WriteLine($"[SPOManager] FedAuth: {(string.IsNullOrEmpty(fedAuth) ? "MISSING" : $"found ({fedAuth.Length} chars)")}");
                System.Diagnostics.Debug.WriteLine($"[SPOManager] rtFa: {(string.IsNullOrEmpty(rtFa) ? "MISSING" : $"found ({rtFa.Length} chars)")}");

                if (!string.IsNullOrEmpty(fedAuth) && !string.IsNullOrEmpty(rtFa))
                {
                    _loginComplete = true;

                    // Try to get current user's email via REST API
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] Fetching current user email...");
                    var userEmail = await GetCurrentUserEmailAsync();
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] User email result: '{userEmail}'");

                    CapturedCookies = new AuthCookies
                    {
                        Domain = _domain,
                        FedAuth = fedAuth,
                        RtFa = rtFa,
                        UserEmail = userEmail,
                        CapturedAt = DateTime.UtcNow
                    };

                    System.Diagnostics.Debug.WriteLine($"[SPOManager] Cookies captured successfully. Closing form.");
                    DialogResult = DialogResult.OK;
                    Close();
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] Error checking cookies: {ex.Message}");
        }
    }

    private async Task<string> GetCurrentUserEmailAsync()
    {
        try
        {
            // Use synchronous XHR because ExecuteScriptAsync doesn't await Promises
            var script = @"
(function() {
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', '/_api/web/currentUser', false);
        xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
        xhr.send(null);
        console.log('XHR status:', xhr.status);
        console.log('XHR response:', xhr.responseText);
        if (xhr.status === 200) {
            var data = JSON.parse(xhr.responseText);
            return data.d.Email || data.d.UserPrincipalName || data.d.LoginName || '';
        }
        return 'ERROR:' + xhr.status;
    } catch (e) {
        return 'EXCEPTION:' + e.toString();
    }
})()";

            System.Diagnostics.Debug.WriteLine($"[SPOManager] Executing JavaScript to get current user...");
            var result = await _webView.CoreWebView2.ExecuteScriptAsync(script);
            System.Diagnostics.Debug.WriteLine($"[SPOManager] JavaScript result (raw): {result}");

            // Result is JSON-encoded, so remove quotes
            if (!string.IsNullOrEmpty(result) && result != "null" && result != "\"\"")
            {
                var cleanResult = result.Trim('"');
                System.Diagnostics.Debug.WriteLine($"[SPOManager] JavaScript result (cleaned): {cleanResult}");

                if (cleanResult.StartsWith("ERROR:") || cleanResult.StartsWith("EXCEPTION:"))
                {
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] JavaScript returned error: {cleanResult}");
                    return string.Empty;
                }

                return cleanResult;
            }

            System.Diagnostics.Debug.WriteLine($"[SPOManager] JavaScript returned null or empty");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SPOManager] Exception getting current user: {ex.Message}");
        }

        return string.Empty;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);

        // Clean up WebView2 events
        if (_webView?.CoreWebView2 != null)
        {
            _webView.CoreWebView2.NavigationCompleted -= OnNavigationCompleted;
        }
    }
}

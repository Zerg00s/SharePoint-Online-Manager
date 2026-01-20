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

            // Check if we've been redirected back to the SharePoint site (login complete)
            if (currentUrl.StartsWith(_siteUrl, StringComparison.OrdinalIgnoreCase) ||
                (currentUrl.Contains(_domain) && !currentUrl.Contains("login.microsoftonline.com")))
            {
                // Try to capture cookies
                var cookies = await _webView.CoreWebView2.CookieManager.GetCookiesAsync($"https://{_domain}");

                var fedAuth = cookies.FirstOrDefault(c => c.Name == "FedAuth")?.Value;
                var rtFa = cookies.FirstOrDefault(c => c.Name == "rtFa")?.Value;

                if (!string.IsNullOrEmpty(fedAuth) && !string.IsNullOrEmpty(rtFa))
                {
                    _loginComplete = true;

                    CapturedCookies = new AuthCookies
                    {
                        Domain = _domain,
                        FedAuth = fedAuth,
                        RtFa = rtFa,
                        CapturedAt = DateTime.UtcNow
                    };

                    DialogResult = DialogResult.OK;
                    Close();
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error checking cookies: {ex.Message}");
        }
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

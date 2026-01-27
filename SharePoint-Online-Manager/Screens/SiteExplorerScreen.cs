using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for exploring a SharePoint site - viewing subsites, lists, and libraries.
/// </summary>
public class SiteExplorerScreen : BaseScreen
{
    private Label _siteInfoLabel = null!;
    private TabControl _tabControl = null!;
    private ListView _subsitesListView = null!;
    private ListView _listsListView = null!;
    private Button _refreshButton = null!;
    private Button _drillDownButton = null!;
    private Button _viewPermissionsButton = null!;

    private string _siteUrl = string.Empty;
    private string _siteTitle = string.Empty;
    private Connection _connection = null!;
    private IAuthenticationService _authService = null!;
    private List<SubsiteInfo> _subsites = [];
    private List<ListInfo> _lists = [];

    private string _screenTitle = "Site Explorer";
    public override string ScreenTitle => _screenTitle;

    protected override void OnInitialize()
    {
        _authService = GetRequiredService<IAuthenticationService>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();

        // Header panel with site info and buttons
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 80,
            Padding = new Padding(10)
        };

        _siteInfoLabel = new Label
        {
            Text = "Loading site information...",
            Font = new Font(Font.FontFamily, 10),
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 40,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 0)
        };

        _refreshButton = new Button
        {
            Text = "\U0001F504 Refresh",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat
        };
        _refreshButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _refreshButton.FlatAppearance.BorderSize = 1;
        _refreshButton.Click += RefreshButton_Click;

        _drillDownButton = new Button
        {
            Text = "\U0001F4C2 Open Selected",
            Size = new Size(130, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _drillDownButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _drillDownButton.FlatAppearance.BorderSize = 1;
        _drillDownButton.Click += DrillDownButton_Click;

        _viewPermissionsButton = new Button
        {
            Text = "\U0001F512 View Permissions",
            Size = new Size(150, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _viewPermissionsButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _viewPermissionsButton.FlatAppearance.BorderSize = 1;
        _viewPermissionsButton.Click += ViewPermissionsButton_Click;

        buttonPanel.Controls.AddRange(new Control[] { _refreshButton, _drillDownButton, _viewPermissionsButton });

        headerPanel.Controls.Add(_siteInfoLabel);
        headerPanel.Controls.Add(buttonPanel);

        // Tab control for subsites and lists
        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        // Subsites tab
        var subsitesTab = new TabPage("Subsites");
        _subsitesListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false
        };
        _subsitesListView.Columns.Add("Title", 200);
        _subsitesListView.Columns.Add("URL", 350);
        _subsitesListView.Columns.Add("Template", 150);
        _subsitesListView.Columns.Add("Created", 120);
        _subsitesListView.Columns.Add("Last Modified", 120);
        _subsitesListView.SelectedIndexChanged += SubsitesListView_SelectedIndexChanged;
        _subsitesListView.DoubleClick += SubsitesListView_DoubleClick;
        subsitesTab.Controls.Add(_subsitesListView);

        // Lists & Libraries tab
        var listsTab = new TabPage("Lists & Libraries");
        _listsListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false
        };
        _listsListView.Columns.Add("Title", 200);
        _listsListView.Columns.Add("Type", 120);
        _listsListView.Columns.Add("Item Count", 80);
        _listsListView.Columns.Add("URL", 300);
        _listsListView.Columns.Add("Created", 120);
        _listsListView.Columns.Add("Last Modified", 120);
        _listsListView.Columns.Add("Hidden", 60);
        _listsListView.SelectedIndexChanged += ListsListView_SelectedIndexChanged;
        listsTab.Controls.Add(_listsListView);

        _tabControl.TabPages.Add(subsitesTab);
        _tabControl.TabPages.Add(listsTab);
        _tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;

        Controls.Add(_tabControl);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        if (parameter is SiteExplorerContext context)
        {
            _siteUrl = context.SiteUrl;
            _connection = context.Connection;
            _siteTitle = context.SiteTitle ?? "Site";
            _screenTitle = $"Explorer - {_siteTitle}";
            UpdateTitle();
        }
        else if (parameter is (string siteUrl, Connection connection))
        {
            _siteUrl = siteUrl;
            _connection = connection;
            _screenTitle = "Site Explorer";
            UpdateTitle();
        }
        else
        {
            MessageBox.Show("No site URL provided.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            await NavigationService!.GoBackAsync();
            return;
        }

        await LoadSiteDataAsync();
    }

    private async Task LoadSiteDataAsync()
    {
        ShowLoading("Loading site information...");

        try
        {
            var cookies = _authService.GetStoredCookies(new Uri(_siteUrl).Host);
            if (cookies == null || !cookies.IsValid)
            {
                var authenticated = await ReauthenticateAsync();
                if (!authenticated)
                {
                    HideLoading();
                    return;
                }
                cookies = _authService.GetStoredCookies(new Uri(_siteUrl).Host);
            }

            if (cookies == null)
            {
                throw new InvalidOperationException("No valid credentials available.");
            }

            using var spService = new SharePointService(cookies, new Uri(_siteUrl).Host);

            // Get site info
            var siteInfo = await spService.GetSiteInfoAsync(_siteUrl);
            _siteTitle = siteInfo.Title;
            _screenTitle = $"Explorer - {_siteTitle}";
            UpdateTitle();

            _siteInfoLabel.Text = $"Title: {siteInfo.Title}\n" +
                                  $"URL: {_siteUrl}\n" +
                                  $"Template: {siteInfo.WebTemplate} | Last Modified: {siteInfo.LastItemModifiedDate:g}";

            // Load subsites
            SetStatus("Loading subsites...");
            _subsites = await spService.GetSubsitesAsync(_siteUrl);
            RefreshSubsitesList();

            // Load lists
            SetStatus("Loading lists and libraries...");
            var listsResult = await spService.GetListsAsync(_siteUrl, includeHidden: true);
            if (listsResult.IsSuccess && listsResult.Data != null)
            {
                _lists = listsResult.Data;
            }
            else
            {
                _lists = [];
            }
            RefreshListsList();

            // Update tab titles with counts
            _tabControl.TabPages[0].Text = $"Subsites ({_subsites.Count})";
            _tabControl.TabPages[1].Text = $"Lists & Libraries ({_lists.Count})";

            SetStatus($"Loaded {_subsites.Count} subsites and {_lists.Count} lists/libraries");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load site data: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Failed to load site data");
        }
        finally
        {
            HideLoading();
        }
    }

    private async Task<bool> ReauthenticateAsync()
    {
        var result = MessageBox.Show(
            "Authentication expired or missing. Would you like to sign in?",
            "Authentication Required",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result != DialogResult.Yes)
        {
            return false;
        }

        using var loginForm = new LoginForm(_siteUrl);
        if (loginForm.ShowDialog(FindForm()) == DialogResult.OK && loginForm.CapturedCookies != null)
        {
            _authService.StoreCookies(loginForm.CapturedCookies);
            return true;
        }

        return false;
    }

    private void RefreshSubsitesList()
    {
        _subsitesListView.Items.Clear();

        foreach (var subsite in _subsites)
        {
            var item = new ListViewItem(subsite.Title)
            {
                Tag = subsite
            };
            item.SubItems.Add(subsite.Url);
            item.SubItems.Add(subsite.WebTemplate);
            item.SubItems.Add(subsite.Created.ToString("g"));
            item.SubItems.Add(subsite.LastItemModifiedDate.ToString("g"));
            _subsitesListView.Items.Add(item);
        }

        UpdateButtonStates();
    }

    private void RefreshListsList()
    {
        _listsListView.Items.Clear();

        foreach (var list in _lists.OrderBy(l => l.Hidden).ThenBy(l => l.Title))
        {
            var listType = GetListTypeName(list.BaseTemplate);
            var item = new ListViewItem(list.Title)
            {
                Tag = list
            };
            item.SubItems.Add(listType);
            item.SubItems.Add(list.ItemCount.ToString());
            item.SubItems.Add(list.ServerRelativeUrl);
            item.SubItems.Add(list.Created.ToString("g"));
            item.SubItems.Add(list.LastItemModifiedDate.ToString("g"));
            item.SubItems.Add(list.Hidden ? "Yes" : "No");

            if (list.Hidden)
            {
                item.ForeColor = Color.Gray;
            }

            _listsListView.Items.Add(item);
        }

        UpdateButtonStates();
    }

    private static string GetListTypeName(int baseTemplate)
    {
        return baseTemplate switch
        {
            100 => "Generic List",
            101 => "Document Library",
            102 => "Survey",
            103 => "Links",
            104 => "Announcements",
            105 => "Contacts",
            106 => "Events",
            107 => "Tasks",
            108 => "Discussion Board",
            109 => "Picture Library",
            110 => "Data Sources",
            115 => "Form Library",
            118 => "Wiki Page Library",
            119 => "Site Pages",
            120 => "Custom List",
            140 => "Workflow History",
            150 => "Promoted Links",
            170 => "Access Requests",
            175 => "App Requests",
            544 => "MicroFeed",
            851 => "Asset Library",
            _ => $"List ({baseTemplate})"
        };
    }

    private void UpdateButtonStates()
    {
        var activeTab = _tabControl.SelectedIndex;

        if (activeTab == 0) // Subsites
        {
            var hasSelection = _subsitesListView.SelectedItems.Count > 0;
            _drillDownButton.Enabled = hasSelection;
            _drillDownButton.Text = "\U0001F4C2 Open Subsite";
            _viewPermissionsButton.Enabled = hasSelection;
        }
        else // Lists
        {
            var hasSelection = _listsListView.SelectedItems.Count > 0;
            _drillDownButton.Enabled = false; // Can't drill into lists
            _drillDownButton.Text = "\U0001F4C2 Open Selected";
            _viewPermissionsButton.Enabled = hasSelection;
        }
    }

    private void TabControl_SelectedIndexChanged(object? sender, EventArgs e)
    {
        UpdateButtonStates();
    }

    private void SubsitesListView_SelectedIndexChanged(object? sender, EventArgs e)
    {
        UpdateButtonStates();
    }

    private void ListsListView_SelectedIndexChanged(object? sender, EventArgs e)
    {
        UpdateButtonStates();
    }

    private async void SubsitesListView_DoubleClick(object? sender, EventArgs e)
    {
        await DrillIntoSelectedSubsiteAsync();
    }

    private async void RefreshButton_Click(object? sender, EventArgs e)
    {
        await LoadSiteDataAsync();
    }

    private async void DrillDownButton_Click(object? sender, EventArgs e)
    {
        await DrillIntoSelectedSubsiteAsync();
    }

    private async Task DrillIntoSelectedSubsiteAsync()
    {
        if (_subsitesListView.SelectedItems.Count == 0)
            return;

        var selectedSubsite = (SubsiteInfo)_subsitesListView.SelectedItems[0].Tag;

        var context = new SiteExplorerContext
        {
            SiteUrl = selectedSubsite.Url,
            SiteTitle = selectedSubsite.Title,
            Connection = _connection
        };

        await NavigationService!.NavigateToAsync<SiteExplorerScreen>(context);
    }

    private void ViewPermissionsButton_Click(object? sender, EventArgs e)
    {
        // TODO: Navigate to permissions view for selected item
        MessageBox.Show("Permission viewing will be implemented in the Permission Report feature.",
            "View Permissions", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}

/// <summary>
/// Context for navigating to the Site Explorer screen.
/// </summary>
public class SiteExplorerContext
{
    public string SiteUrl { get; set; } = string.Empty;
    public string? SiteTitle { get; set; }
    public Connection Connection { get; set; } = null!;
}

/// <summary>
/// Represents a subsite/subweb in SharePoint.
/// </summary>
public class SubsiteInfo
{
    public string Title { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
    public string ServerRelativeUrl { get; set; } = string.Empty;
    public string WebTemplate { get; set; } = string.Empty;
    public DateTime Created { get; set; }
    public DateTime LastItemModifiedDate { get; set; }
}

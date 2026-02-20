using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Home screen displaying the list of saved connections.
/// </summary>
public class HomeScreen : BaseScreen
{
    private ListView _connectionsListView = null!;
    private ListView _tenantPairsListView = null!;
    private Button _newAdminConnectionButton = null!;
    private Button _newSiteConnectionButton = null!;
    private Button _connectButton = null!;
    private Button _reauthButton = null!;
    private Button _deleteButton = null!;
    private Button _setRoleButton = null!;
    private Button _listCompareButton = null!;
    private Button _navSettingsButton = null!;
    private Button _docCompareButton = null!;
    private Button _siteAccessButton = null!;
    private Button _addPairButton = null!;
    private Button _deletePairButton = null!;
    private ToolTip _toolTip = null!;
    private IConnectionManager _connectionManager = null!;
    private IAuthenticationService _authService = null!;
    private ITenantPairService _tenantPairService = null!;

    public override string ScreenTitle => "Connections";
    public override bool ShowBackButton => false;

    protected override void OnInitialize()
    {
        _connectionManager = GetRequiredService<IConnectionManager>();
        _authService = GetRequiredService<IAuthenticationService>();
        _tenantPairService = GetRequiredService<ITenantPairService>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();

        // Initialize tooltip
        _toolTip = new ToolTip
        {
            AutoPopDelay = 5000,
            InitialDelay = 500,
            ReshowDelay = 100,
            ShowAlways = true
        };

        // Header panel with buttons
        var headerPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 45,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        _newAdminConnectionButton = new Button
        {
            Text = "\u2795 New Admin Connection", // Plus sign
            Size = new Size(170, 32),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _newAdminConnectionButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _newAdminConnectionButton.FlatAppearance.BorderSize = 1;
        _newAdminConnectionButton.Click += NewAdminConnectionButton_Click;

        _newSiteConnectionButton = new Button
        {
            Text = "\u2795 New Site Connection", // Plus sign
            Size = new Size(160, 32),
            Margin = new Padding(0, 0, 20, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _newSiteConnectionButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _newSiteConnectionButton.FlatAppearance.BorderSize = 1;
        _newSiteConnectionButton.Click += NewSiteConnectionButton_Click;

        _connectButton = new Button
        {
            Text = "\U0001F517 Connect", // Link emoji
            Size = new Size(100, 32),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _connectButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _connectButton.FlatAppearance.BorderSize = 1;
        _connectButton.Click += ConnectButton_Click;

        _reauthButton = new Button
        {
            Text = "\U0001F511 Re-authenticate", // Key emoji
            Size = new Size(130, 32),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _reauthButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _reauthButton.FlatAppearance.BorderSize = 1;
        _reauthButton.Click += ReauthButton_Click;

        _deleteButton = new Button
        {
            Text = "\U0001F5D1 Delete", // Wastebasket emoji
            Size = new Size(90, 32),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat,
            ForeColor = Color.DarkRed
        };
        _deleteButton.FlatAppearance.BorderColor = Color.DarkRed;
        _deleteButton.FlatAppearance.BorderSize = 1;
        _deleteButton.Click += DeleteButton_Click;

        _setRoleButton = new Button
        {
            Text = "\U0001F3AF Set Role", // Target emoji
            Size = new Size(100, 32),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _setRoleButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _setRoleButton.FlatAppearance.BorderSize = 1;
        _setRoleButton.Click += SetRoleButton_Click;

        _listCompareButton = new Button
        {
            Text = "\U0001F504 Compare Lists", // Arrows clockwise emoji
            Size = new Size(130, 32),
            Margin = new Padding(20, 0, 10, 0),
            FlatStyle = FlatStyle.Flat
        };
        _listCompareButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _listCompareButton.FlatAppearance.BorderSize = 1;
        _listCompareButton.Click += ListCompareButton_Click;

        _navSettingsButton = new Button
        {
            Text = "\U0001F517 Nav Settings Sync", // Link emoji
            Size = new Size(150, 32),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat
        };
        _navSettingsButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _navSettingsButton.FlatAppearance.BorderSize = 1;
        _navSettingsButton.Click += NavSettingsButton_Click;

        _docCompareButton = new Button
        {
            Text = "\U0001F4C4 Compare Documents", // Page emoji
            Size = new Size(160, 32),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat
        };
        _docCompareButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _docCompareButton.FlatAppearance.BorderSize = 1;
        _docCompareButton.Click += DocCompareButton_Click;

        _siteAccessButton = new Button
        {
            Text = "\U0001F511 Site Access Check", // Key emoji
            Size = new Size(150, 32),
            Margin = new Padding(0, 0, 0, 0),
            FlatStyle = FlatStyle.Flat
        };
        _siteAccessButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _siteAccessButton.FlatAppearance.BorderSize = 1;
        _siteAccessButton.Click += SiteAccessButton_Click;

        // Add tooltips to task buttons
        _toolTip.SetToolTip(_listCompareButton, "Compare list item counts between source and target sites");
        _toolTip.SetToolTip(_navSettingsButton, "Compare and sync navigation settings between tenants");
        _toolTip.SetToolTip(_docCompareButton, "Compare documents between source and target sites");
        _toolTip.SetToolTip(_siteAccessButton, "Check site access for source and target accounts");
        _toolTip.SetToolTip(_newAdminConnectionButton, "Create a new admin-level connection to a SharePoint tenant");
        _toolTip.SetToolTip(_newSiteConnectionButton, "Create a new connection to a specific site collection");
        _toolTip.SetToolTip(_connectButton, "Connect to the selected SharePoint tenant or site");
        _toolTip.SetToolTip(_reauthButton, "Clear cached credentials and sign in again");
        _toolTip.SetToolTip(_deleteButton, "Delete the selected connection");
        _toolTip.SetToolTip(_setRoleButton, "Set the role (Source/Target) for the selected connection");

        headerPanel.Controls.AddRange(new Control[]
        {
            _newAdminConnectionButton,
            _newSiteConnectionButton,
            _connectButton,
            _reauthButton,
            _deleteButton,
            _setRoleButton,
            _listCompareButton,
            _navSettingsButton,
            _docCompareButton,
            _siteAccessButton
        });

        // Main split container for connections and tenant pairs
        var splitContainer = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterDistance = 250,
            Panel1MinSize = 100,
            Panel2MinSize = 100
        };

        // Connections ListView
        _connectionsListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false
        };
        _connectionsListView.Columns.Add("Name", 120);
        _connectionsListView.Columns.Add("Role", 55);
        _connectionsListView.Columns.Add("Type", 55);
        _connectionsListView.Columns.Add("Tenant/URL", 140);
        _connectionsListView.Columns.Add("Account", 140);
        _connectionsListView.Columns.Add("Token Expires", 140);
        _connectionsListView.Columns.Add("Token Life", 65);
        _connectionsListView.Columns.Add("Last Connected", 100);

        _connectionsListView.SelectedIndexChanged += ConnectionsListView_SelectedIndexChanged;
        _connectionsListView.DoubleClick += ConnectionsListView_DoubleClick;

        splitContainer.Panel1.Controls.Add(_connectionsListView);

        // Tenant Pairs section
        var pairsPanel = new Panel { Dock = DockStyle.Fill };

        var pairsHeaderPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 40,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        var pairsLabel = new Label
        {
            Text = "Tenant Pairs (Source \u2192 Target)",
            Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
            AutoSize = true,
            Margin = new Padding(0, 6, 20, 0)
        };

        _addPairButton = new Button
        {
            Text = "\u2795 Add Pair",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _addPairButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _addPairButton.FlatAppearance.BorderSize = 1;
        _addPairButton.Click += AddPairButton_Click;

        _deletePairButton = new Button
        {
            Text = "\U0001F5D1 Delete Pair",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat,
            ForeColor = Color.DarkRed
        };
        _deletePairButton.FlatAppearance.BorderColor = Color.DarkRed;
        _deletePairButton.FlatAppearance.BorderSize = 1;
        _deletePairButton.Click += DeletePairButton_Click;

        _toolTip.SetToolTip(_addPairButton, "Add a new source-to-target tenant pair");
        _toolTip.SetToolTip(_deletePairButton, "Delete the selected tenant pair");

        pairsHeaderPanel.Controls.AddRange(new Control[] { pairsLabel, _addPairButton, _deletePairButton });

        _tenantPairsListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false
        };
        _tenantPairsListView.Columns.Add("Source Tenant", 200);
        _tenantPairsListView.Columns.Add("", 40);
        _tenantPairsListView.Columns.Add("Target Tenant", 200);
        _tenantPairsListView.Columns.Add("Name", 150);
        _tenantPairsListView.SelectedIndexChanged += TenantPairsListView_SelectedIndexChanged;
        _tenantPairsListView.DoubleClick += TenantPairsListView_DoubleClick;

        pairsPanel.Controls.Add(_tenantPairsListView);
        pairsPanel.Controls.Add(pairsHeaderPanel);

        splitContainer.Panel2.Controls.Add(pairsPanel);

        Controls.Add(splitContainer);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        await RefreshConnectionsAsync();
        await RefreshTenantPairsAsync();
    }

    private async Task RefreshConnectionsAsync()
    {
        ShowLoading("Loading connections...");

        try
        {
            var connections = await _connectionManager.GetAllConnectionsAsync();
            UpdateConnectionsList(connections);
        }
        finally
        {
            HideLoading();
        }
    }

    private void UpdateConnectionsList(List<Connection> connections)
    {
        _connectionsListView.Items.Clear();

        foreach (var conn in connections)
        {
            var item = new ListViewItem(conn.Name)
            {
                Tag = conn
            };
            item.SubItems.Add(conn.RoleDescription);
            item.SubItems.Add(conn.Type.ToString());
            item.SubItems.Add(conn.Type == ConnectionType.Admin ? conn.TenantName : conn.SiteUrl ?? "");

            var hasCredentials = _connectionManager.HasStoredCredentials(conn);
            var credentialDisplay = "None";
            var expirationDisplay = "-";
            var tokenLifeDisplay = "-";

            if (hasCredentials)
            {
                // Try to get the user email and expiration from stored cookies
                // Check both admin domain and tenant domain for Admin connections
                AuthCookies? cookies = null;
                if (conn.Type == ConnectionType.Admin)
                {
                    cookies = _authService.GetStoredCookies(conn.AdminDomain)
                              ?? _authService.GetStoredCookies(conn.TenantDomain);
                }
                else if (!string.IsNullOrEmpty(conn.SiteUrl))
                {
                    cookies = _authService.GetStoredCookies(new Uri(conn.SiteUrl).Host);
                }

                if (cookies != null)
                {
                    credentialDisplay = !string.IsNullOrEmpty(cookies.UserEmail) ? cookies.UserEmail : "Stored";
                    expirationDisplay = cookies.ExpirationDisplay;
                    tokenLifeDisplay = cookies.TotalDurationDisplay;

                    // Color code based on expiration status
                    if (cookies.IsExpired)
                    {
                        item.BackColor = Color.FromArgb(255, 200, 200); // Red - expired
                    }
                    else if (cookies.TimeRemaining.HasValue && cookies.TimeRemaining.Value.TotalHours < 1)
                    {
                        item.BackColor = Color.FromArgb(255, 255, 200); // Yellow - expiring soon
                    }
                    else
                    {
                        item.BackColor = Color.FromArgb(230, 255, 230); // Green - valid
                    }
                }
                else
                {
                    credentialDisplay = "Stored";
                    item.BackColor = Color.FromArgb(230, 255, 230);
                }
            }

            item.SubItems.Add(credentialDisplay);
            item.SubItems.Add(expirationDisplay);
            item.SubItems.Add(tokenLifeDisplay);
            item.SubItems.Add(conn.LastConnectedAt?.ToString("g") ?? "Never");

            _connectionsListView.Items.Add(item);
        }

        UpdateButtonStates();
    }

    private void ConnectionsListView_SelectedIndexChanged(object? sender, EventArgs e)
    {
        UpdateButtonStates();
    }

    private void UpdateButtonStates()
    {
        var hasSelection = _connectionsListView.SelectedItems.Count > 0;
        _connectButton.Enabled = hasSelection;
        _reauthButton.Enabled = hasSelection;
        _deleteButton.Enabled = hasSelection;
        _setRoleButton.Enabled = hasSelection;
    }

    private async void ConnectionsListView_DoubleClick(object? sender, EventArgs e)
    {
        await OpenSelectedConnectionAsync();
    }

    private async void ConnectButton_Click(object? sender, EventArgs e)
    {
        await OpenSelectedConnectionAsync();
    }

    private async Task OpenSelectedConnectionAsync()
    {
        if (_connectionsListView.SelectedItems.Count == 0)
            return;

        var connection = (Connection)_connectionsListView.SelectedItems[0].Tag;

        // Check if we need to authenticate
        if (!_connectionManager.HasStoredCredentials(connection))
        {
            var authenticated = await AuthenticateAsync(connection);
            if (!authenticated)
            {
                return;
            }
        }

        // Update last connected
        await _connectionManager.UpdateLastConnectedAsync(connection.Id);

        // Navigate to site collections screen
        await NavigationService!.NavigateToAsync<SiteCollectionsScreen>(connection);
    }

    private async Task<bool> AuthenticateAsync(Connection connection)
    {
        SetStatus("Opening sign-in window...");

        try
        {
            using var loginForm = new LoginForm(connection.PrimaryUrl);
            var result = loginForm.ShowDialog(FindForm());

            if (result == DialogResult.OK && loginForm.CapturedCookies != null)
            {
                _authService.StoreCookies(loginForm.CapturedCookies);
                SetStatus($"Authenticated to {connection.CookieDomain}");
                await RefreshConnectionsAsync();
                return true;
            }
            else
            {
                SetStatus("Authentication cancelled");
                return false;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Authentication error: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    private async void NewAdminConnectionButton_Click(object? sender, EventArgs e)
    {
        await CreateNewConnectionAsync(ConnectionType.Admin);
    }

    private async void NewSiteConnectionButton_Click(object? sender, EventArgs e)
    {
        await CreateNewConnectionAsync(ConnectionType.SiteCollection);
    }

    private async Task CreateNewConnectionAsync(ConnectionType type)
    {
        using var dialog = new AddConnectionDialog(type);
        if (dialog.ShowDialog(FindForm()) == DialogResult.OK && dialog.Connection != null)
        {
            await _connectionManager.SaveConnectionAsync(dialog.Connection);
            await RefreshConnectionsAsync();
            SetStatus($"Connection '{dialog.Connection.Name}' created");
        }
    }

    private async void ReauthButton_Click(object? sender, EventArgs e)
    {
        if (_connectionsListView.SelectedItems.Count == 0)
            return;

        var connection = (Connection)_connectionsListView.SelectedItems[0].Tag;

        // Clear stored credentials for this connection
        if (connection.Type == ConnectionType.Admin)
        {
            _authService.ClearCredentials(connection.AdminDomain);
            _authService.ClearCredentials(connection.TenantDomain);
        }
        else if (!string.IsNullOrEmpty(connection.SiteUrl))
        {
            _authService.ClearCredentials(new Uri(connection.SiteUrl).Host);
        }

        // Clear the WebView2 cache to force fresh login
        try
        {
            var webView2Folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SharePointOnlineManager", "WebView2");

            if (Directory.Exists(webView2Folder))
            {
                Directory.Delete(webView2Folder, true);
                SetStatus("Browser cache cleared");
            }
        }
        catch (Exception ex)
        {
            SetStatus($"Warning: Could not clear browser cache: {ex.Message}");
        }

        // Now authenticate fresh
        var authenticated = await AuthenticateAsync(connection);
        if (authenticated)
        {
            connection.LastConnectedAt = DateTime.UtcNow;
            await _connectionManager.SaveConnectionAsync(connection);
            await RefreshConnectionsAsync();
            SetStatus($"Re-authenticated to '{connection.Name}'");
        }
    }

    private async void DeleteButton_Click(object? sender, EventArgs e)
    {
        if (_connectionsListView.SelectedItems.Count == 0)
            return;

        var connection = (Connection)_connectionsListView.SelectedItems[0].Tag;

        var result = MessageBox.Show(
            $"Are you sure you want to delete the connection '{connection.Name}'?\n\nThis will also remove any stored credentials.",
            "Delete Connection",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            await _connectionManager.DeleteConnectionAsync(connection.Id);
            await RefreshConnectionsAsync();
            SetStatus($"Connection '{connection.Name}' deleted");
        }
    }

    private async void ListCompareButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.NavigateToAsync<ListCompareConfigScreen>();
    }

    private async void NavSettingsButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.NavigateToAsync<NavigationSettingsConfigScreen>();
    }

    private async void DocCompareButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.NavigateToAsync<DocumentCompareConfigScreen>();
    }

    private async void SiteAccessButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.NavigateToAsync<SiteAccessConfigScreen>();
    }

    private async void SetRoleButton_Click(object? sender, EventArgs e)
    {
        if (_connectionsListView.SelectedItems.Count == 0)
            return;

        var connection = (Connection)_connectionsListView.SelectedItems[0].Tag;

        using var dialog = new SetRoleDialog(connection.Role);
        if (dialog.ShowDialog(FindForm()) == DialogResult.OK)
        {
            connection.Role = dialog.SelectedRole;
            await _connectionManager.SaveConnectionAsync(connection);
            await RefreshConnectionsAsync();
            SetStatus($"Role set to '{connection.RoleDescription}' for '{connection.Name}'");
        }
    }

    private void TenantPairsListView_SelectedIndexChanged(object? sender, EventArgs e)
    {
        _deletePairButton.Enabled = _tenantPairsListView.SelectedItems.Count > 0;
    }

    private async void TenantPairsListView_DoubleClick(object? sender, EventArgs e)
    {
        if (_tenantPairsListView.SelectedItems.Count == 0)
            return;

        var pair = (TenantPair)_tenantPairsListView.SelectedItems[0].Tag;
        await NavigationService!.NavigateToAsync<TenantPairDetailScreen>(pair);
    }

    private async void AddPairButton_Click(object? sender, EventArgs e)
    {
        var connections = await _connectionManager.GetAllConnectionsAsync();
        var sourceConnections = connections.Where(c => c.Role == TenantRole.Source).ToList();
        var targetConnections = connections.Where(c => c.Role == TenantRole.Target).ToList();

        if (sourceConnections.Count == 0 || targetConnections.Count == 0)
        {
            MessageBox.Show(
                "You need at least one Source connection and one Target connection to create a pair.\n\n" +
                "Use the 'Set Role' button to mark connections as Source or Target.",
                "Cannot Create Pair",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        using var dialog = new AddTenantPairDialog(sourceConnections, targetConnections);
        if (dialog.ShowDialog(FindForm()) == DialogResult.OK && dialog.TenantPair != null)
        {
            await _tenantPairService.SavePairAsync(dialog.TenantPair);
            await RefreshTenantPairsAsync();
            SetStatus("Tenant pair added");
        }
    }

    private async void DeletePairButton_Click(object? sender, EventArgs e)
    {
        if (_tenantPairsListView.SelectedItems.Count == 0)
            return;

        var pair = (TenantPair)_tenantPairsListView.SelectedItems[0].Tag;

        var result = MessageBox.Show(
            "Are you sure you want to delete this tenant pair?",
            "Delete Pair",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            await _tenantPairService.DeletePairAsync(pair.Id);
            await RefreshTenantPairsAsync();
            SetStatus("Tenant pair deleted");
        }
    }

    private async Task RefreshTenantPairsAsync()
    {
        var pairs = await _tenantPairService.GetAllPairsAsync();
        var connections = await _connectionManager.GetAllConnectionsAsync();

        _tenantPairsListView.Items.Clear();

        foreach (var pair in pairs)
        {
            var sourceConn = connections.FirstOrDefault(c => c.Id == pair.SourceConnectionId);
            var targetConn = connections.FirstOrDefault(c => c.Id == pair.TargetConnectionId);

            var item = new ListViewItem(sourceConn?.Name ?? "(deleted)")
            {
                Tag = pair
            };
            item.SubItems.Add("\u2192"); // Arrow
            item.SubItems.Add(targetConn?.Name ?? "(deleted)");
            item.SubItems.Add(pair.Name ?? "");

            // Gray out if connections are missing
            if (sourceConn == null || targetConn == null)
            {
                item.ForeColor = Color.Gray;
            }

            _tenantPairsListView.Items.Add(item);
        }

        _deletePairButton.Enabled = false;
    }
}

/// <summary>
/// Dialog for adding a new connection.
/// </summary>
public class AddConnectionDialog : Form
{
    private TextBox _nameTextBox = null!;
    private TextBox _tenantTextBox = null!;
    private TextBox _siteUrlTextBox = null!;
    private Label _siteUrlLabel = null!;
    private readonly ConnectionType _connectionType;

    public Connection? Connection { get; private set; }

    public AddConnectionDialog(ConnectionType type)
    {
        _connectionType = type;
        InitializeUI();
    }

    private void InitializeUI()
    {
        Text = _connectionType == ConnectionType.Admin
            ? "New Admin Connection"
            : "New Site Collection Connection";
        Size = new Size(450, 220);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;

        var nameLabel = new Label
        {
            Text = "Connection Name:",
            Location = new Point(15, 20),
            AutoSize = true
        };

        _nameTextBox = new TextBox
        {
            Location = new Point(15, 40),
            Size = new Size(400, 23)
        };

        var tenantLabel = new Label
        {
            Text = "Tenant Name (e.g., 'contoso'):",
            Location = new Point(15, 70),
            AutoSize = true
        };

        _tenantTextBox = new TextBox
        {
            Location = new Point(15, 90),
            Size = new Size(200, 23)
        };

        _siteUrlLabel = new Label
        {
            Text = "Site URL:",
            Location = new Point(15, 70),
            AutoSize = true,
            Visible = _connectionType == ConnectionType.SiteCollection
        };

        _siteUrlTextBox = new TextBox
        {
            Location = new Point(15, 90),
            Size = new Size(400, 23),
            Visible = _connectionType == ConnectionType.SiteCollection
        };

        if (_connectionType == ConnectionType.SiteCollection)
        {
            tenantLabel.Visible = false;
            _tenantTextBox.Visible = false;
        }

        var okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Location = new Point(255, 140),
            Size = new Size(75, 28)
        };
        okButton.Click += OkButton_Click;

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Location = new Point(340, 140),
            Size = new Size(75, 28)
        };

        AcceptButton = okButton;
        CancelButton = cancelButton;

        Controls.AddRange(new Control[]
        {
            nameLabel, _nameTextBox,
            tenantLabel, _tenantTextBox,
            _siteUrlLabel, _siteUrlTextBox,
            okButton, cancelButton
        });
    }

    private void OkButton_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_nameTextBox.Text))
        {
            MessageBox.Show("Please enter a connection name.", "Validation",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.None;
            return;
        }

        if (_connectionType == ConnectionType.Admin)
        {
            if (string.IsNullOrWhiteSpace(_tenantTextBox.Text))
            {
                MessageBox.Show("Please enter a tenant name.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            Connection = new Connection
            {
                Name = _nameTextBox.Text.Trim(),
                Type = ConnectionType.Admin,
                TenantName = _tenantTextBox.Text.Trim().ToLowerInvariant()
            };
        }
        else
        {
            if (string.IsNullOrWhiteSpace(_siteUrlTextBox.Text) ||
                !Uri.TryCreate(_siteUrlTextBox.Text, UriKind.Absolute, out var uri))
            {
                MessageBox.Show("Please enter a valid site URL.", "Validation",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            // Extract tenant name from URL
            var host = uri.Host;
            var tenantName = host.Split('.')[0];
            if (tenantName.EndsWith("-admin"))
            {
                tenantName = tenantName[..^6];
            }
            if (tenantName.EndsWith("-my"))
            {
                tenantName = tenantName[..^3];
            }

            Connection = new Connection
            {
                Name = _nameTextBox.Text.Trim(),
                Type = ConnectionType.SiteCollection,
                TenantName = tenantName,
                SiteUrl = _siteUrlTextBox.Text.Trim()
            };
        }
    }
}

/// <summary>
/// Dialog for setting the role of a connection.
/// </summary>
public class SetRoleDialog : Form
{
    private ComboBox _roleComboBox = null!;

    public TenantRole SelectedRole { get; private set; }

    public SetRoleDialog(TenantRole currentRole)
    {
        SelectedRole = currentRole;
        InitializeUI();
    }

    private void InitializeUI()
    {
        Text = "Set Connection Role";
        Size = new Size(300, 150);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;

        var roleLabel = new Label
        {
            Text = "Role:",
            Location = new Point(15, 20),
            AutoSize = true
        };

        _roleComboBox = new ComboBox
        {
            Location = new Point(15, 45),
            Size = new Size(250, 23),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _roleComboBox.Items.AddRange(["(None)", "Source", "Target"]);
        _roleComboBox.SelectedIndex = SelectedRole switch
        {
            TenantRole.Source => 1,
            TenantRole.Target => 2,
            _ => 0
        };

        var okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Location = new Point(105, 80),
            Size = new Size(75, 28)
        };
        okButton.Click += (s, e) =>
        {
            SelectedRole = _roleComboBox.SelectedIndex switch
            {
                1 => TenantRole.Source,
                2 => TenantRole.Target,
                _ => TenantRole.Unspecified
            };
        };

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Location = new Point(190, 80),
            Size = new Size(75, 28)
        };

        AcceptButton = okButton;
        CancelButton = cancelButton;

        Controls.AddRange(new Control[] { roleLabel, _roleComboBox, okButton, cancelButton });
    }
}

/// <summary>
/// Dialog for adding a tenant pair.
/// </summary>
public class AddTenantPairDialog : Form
{
    private ComboBox _sourceComboBox = null!;
    private ComboBox _targetComboBox = null!;
    private TextBox _nameTextBox = null!;
    private readonly List<Connection> _sourceConnections;
    private readonly List<Connection> _targetConnections;

    public TenantPair? TenantPair { get; private set; }

    public AddTenantPairDialog(List<Connection> sourceConnections, List<Connection> targetConnections)
    {
        _sourceConnections = sourceConnections;
        _targetConnections = targetConnections;
        InitializeUI();
    }

    private void InitializeUI()
    {
        Text = "Add Tenant Pair";
        Size = new Size(400, 260);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;

        var sourceLabel = new Label
        {
            Text = "Source Tenant:",
            Location = new Point(15, 20),
            AutoSize = true
        };

        _sourceComboBox = new ComboBox
        {
            Location = new Point(15, 40),
            Size = new Size(350, 23),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        foreach (var conn in _sourceConnections)
        {
            _sourceComboBox.Items.Add(conn.Name);
        }
        if (_sourceComboBox.Items.Count > 0)
            _sourceComboBox.SelectedIndex = 0;

        var targetLabel = new Label
        {
            Text = "Target Tenant:",
            Location = new Point(15, 70),
            AutoSize = true
        };

        _targetComboBox = new ComboBox
        {
            Location = new Point(15, 90),
            Size = new Size(350, 23),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        foreach (var conn in _targetConnections)
        {
            _targetComboBox.Items.Add(conn.Name);
        }
        if (_targetComboBox.Items.Count > 0)
            _targetComboBox.SelectedIndex = 0;

        var nameLabel = new Label
        {
            Text = "Pair Name (optional):",
            Location = new Point(15, 120),
            AutoSize = true
        };

        _nameTextBox = new TextBox
        {
            Location = new Point(15, 140),
            Size = new Size(350, 23)
        };

        var okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Location = new Point(205, 175),
            Size = new Size(75, 28)
        };
        okButton.Click += OkButton_Click;

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Location = new Point(290, 175),
            Size = new Size(75, 28)
        };

        AcceptButton = okButton;
        CancelButton = cancelButton;

        Controls.AddRange(new Control[]
        {
            sourceLabel, _sourceComboBox,
            targetLabel, _targetComboBox,
            nameLabel, _nameTextBox,
            okButton, cancelButton
        });
    }

    private void OkButton_Click(object? sender, EventArgs e)
    {
        if (_sourceComboBox.SelectedIndex < 0 || _targetComboBox.SelectedIndex < 0)
        {
            MessageBox.Show("Please select both source and target connections.", "Validation",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            DialogResult = DialogResult.None;
            return;
        }

        TenantPair = new TenantPair
        {
            SourceConnectionId = _sourceConnections[_sourceComboBox.SelectedIndex].Id,
            TargetConnectionId = _targetConnections[_targetComboBox.SelectedIndex].Id,
            Name = string.IsNullOrWhiteSpace(_nameTextBox.Text) ? null : _nameTextBox.Text.Trim()
        };
    }
}

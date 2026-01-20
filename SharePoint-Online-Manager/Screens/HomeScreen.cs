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
    private Button _newAdminConnectionButton = null!;
    private Button _newSiteConnectionButton = null!;
    private Button _connectButton = null!;
    private Button _deleteButton = null!;
    private Button _listCompareButton = null!;
    private IConnectionManager _connectionManager = null!;
    private IAuthenticationService _authService = null!;

    public override string ScreenTitle => "Connections";
    public override bool ShowBackButton => false;

    protected override void OnInitialize()
    {
        _connectionManager = GetRequiredService<IConnectionManager>();
        _authService = GetRequiredService<IAuthenticationService>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();

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
            Text = "New Admin Connection",
            Size = new Size(150, 32),
            Margin = new Padding(0, 0, 10, 0)
        };
        _newAdminConnectionButton.Click += NewAdminConnectionButton_Click;

        _newSiteConnectionButton = new Button
        {
            Text = "New Site Connection",
            Size = new Size(150, 32),
            Margin = new Padding(0, 0, 20, 0)
        };
        _newSiteConnectionButton.Click += NewSiteConnectionButton_Click;

        _connectButton = new Button
        {
            Text = "Connect",
            Size = new Size(100, 32),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false
        };
        _connectButton.Click += ConnectButton_Click;

        _deleteButton = new Button
        {
            Text = "Delete",
            Size = new Size(80, 32),
            Enabled = false
        };
        _deleteButton.Click += DeleteButton_Click;

        _listCompareButton = new Button
        {
            Text = "List Compare",
            Size = new Size(110, 32),
            Margin = new Padding(20, 0, 0, 0)
        };
        _listCompareButton.Click += ListCompareButton_Click;

        headerPanel.Controls.AddRange(new Control[]
        {
            _newAdminConnectionButton,
            _newSiteConnectionButton,
            _connectButton,
            _deleteButton,
            _listCompareButton
        });

        // Connections ListView
        _connectionsListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false
        };
        _connectionsListView.Columns.Add("Name", 200);
        _connectionsListView.Columns.Add("Type", 120);
        _connectionsListView.Columns.Add("Tenant/URL", 300);
        _connectionsListView.Columns.Add("Last Connected", 150);
        _connectionsListView.Columns.Add("Credentials", 100);

        _connectionsListView.SelectedIndexChanged += ConnectionsListView_SelectedIndexChanged;
        _connectionsListView.DoubleClick += ConnectionsListView_DoubleClick;

        Controls.Add(_connectionsListView);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        await RefreshConnectionsAsync();
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
            item.SubItems.Add(conn.Type.ToString());
            item.SubItems.Add(conn.Type == ConnectionType.Admin ? conn.TenantName : conn.SiteUrl ?? "");
            item.SubItems.Add(conn.LastConnectedAt?.ToString("g") ?? "Never");

            var hasCredentials = _connectionManager.HasStoredCredentials(conn);
            item.SubItems.Add(hasCredentials ? "Stored" : "None");

            if (hasCredentials)
            {
                item.BackColor = Color.FromArgb(230, 255, 230);
            }

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
        _deleteButton.Enabled = hasSelection;
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

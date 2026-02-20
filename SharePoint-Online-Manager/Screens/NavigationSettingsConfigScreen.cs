using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for configuring a new Navigation Settings Sync task.
/// </summary>
public class NavigationSettingsConfigScreen : BaseScreen
{
    private ComboBox _sourceConnectionCombo = null!;
    private ComboBox _targetConnectionCombo = null!;
    private Button _importCsvButton = null!;
    private Button _generateTargetUrlsButton = null!;
    private DataGridView _sitePairsGrid = null!;
    private TextBox _taskNameTextBox = null!;
    private Button _createTaskButton = null!;
    private Button _clearPairsButton = null!;
    private Label _selectedSitesLabel = null!;

    private IConnectionManager _connectionManager = null!;
    private IAuthenticationService _authService = null!;
    private ITaskService _taskService = null!;
    private ITenantPairService _tenantPairService = null!;
    private List<Connection> _connections = [];
    private List<SiteComparePair> _sitePairs = [];
    private TaskCreationContext? _context;
    private List<string> _selectedSourceUrls = [];

    public override string ScreenTitle => "Create Navigation Settings Sync Task";

    protected override void OnInitialize()
    {
        _connectionManager = GetRequiredService<IConnectionManager>();
        _authService = GetRequiredService<IAuthenticationService>();
        _taskService = GetRequiredService<ITaskService>();
        _tenantPairService = GetRequiredService<ITenantPairService>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();
        AutoScroll = true;

        var mainPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 6,
            Padding = new Padding(10)
        };
        mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 70)); // Connections row
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45)); // Selected sites info row
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45)); // Buttons row
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Site pairs grid
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60)); // Task name
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45)); // Create button

        // Row 0: Connection selectors
        var sourceConnPanel = CreateConnectionPanel("Source Connection:", out _sourceConnectionCombo);
        var targetConnPanel = CreateConnectionPanel("Target Connection:", out _targetConnectionCombo);
        mainPanel.Controls.Add(sourceConnPanel, 0, 0);
        mainPanel.Controls.Add(targetConnPanel, 1, 0);

        // Row 1: Selected sites info
        var selectedSitesPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        _selectedSitesLabel = new Label
        {
            Text = "No sites selected from Site Collections screen",
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
            Padding = new Padding(0, 8, 0, 0)
        };

        selectedSitesPanel.Controls.Add(_selectedSitesLabel);
        mainPanel.Controls.Add(selectedSitesPanel, 0, 1);
        mainPanel.SetColumnSpan(selectedSitesPanel, 2);

        // Row 2: Buttons row
        var buttonsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        _generateTargetUrlsButton = new Button
        {
            Text = "Generate Target URLs",
            Size = new Size(150, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false
        };
        _generateTargetUrlsButton.Click += GenerateTargetUrlsButton_Click;

        _importCsvButton = new Button
        {
            Text = "Import CSV File",
            Size = new Size(120, 28),
            Margin = new Padding(0, 0, 10, 0)
        };
        _importCsvButton.Click += ImportCsvButton_Click;

        _clearPairsButton = new Button
        {
            Text = "Clear All",
            Size = new Size(80, 28),
            Enabled = false
        };
        _clearPairsButton.Click += ClearPairsButton_Click;

        var csvInfoLabel = new Label
        {
            Text = "CSV format: Source URL, Target URL",
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
            Padding = new Padding(10, 8, 0, 0)
        };

        buttonsPanel.Controls.AddRange(new Control[] { _generateTargetUrlsButton, _importCsvButton, _clearPairsButton, csvInfoLabel });
        mainPanel.Controls.Add(buttonsPanel, 0, 2);
        mainPanel.SetColumnSpan(buttonsPanel, 2);

        // Row 3: Site pairs grid
        var sitePairsPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
        var sitePairsLabel = new Label
        {
            Text = "Site Pairs (root sites only):",
            Dock = DockStyle.Top,
            Height = 20
        };

        _sitePairsGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = true,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };
        _sitePairsGrid.Columns.Add("SourceUrl", "Source URL");
        _sitePairsGrid.Columns.Add("TargetUrl", "Target URL");

        sitePairsPanel.Controls.Add(_sitePairsGrid);
        sitePairsPanel.Controls.Add(sitePairsLabel);
        mainPanel.Controls.Add(sitePairsPanel, 0, 3);
        mainPanel.SetColumnSpan(sitePairsPanel, 2);

        // Row 4: Task name
        var taskNamePanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
        var taskNameLabel = new Label
        {
            Text = "Task Name:",
            Dock = DockStyle.Top,
            Height = 20
        };

        _taskNameTextBox = new TextBox
        {
            Dock = DockStyle.Top,
            Text = $"Navigation Settings Sync - {DateTime.Now:yyyy-MM-dd HH:mm}"
        };

        taskNamePanel.Controls.Add(_taskNameTextBox);
        taskNamePanel.Controls.Add(taskNameLabel);
        mainPanel.Controls.Add(taskNamePanel, 0, 4);
        mainPanel.SetColumnSpan(taskNamePanel, 2);

        // Row 5: Create button
        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(0, 5, 0, 5)
        };

        _createTaskButton = new Button
        {
            Text = "Create Task",
            Size = new Size(120, 32),
            Enabled = false
        };
        _createTaskButton.Click += CreateTaskButton_Click;

        buttonPanel.Controls.Add(_createTaskButton);
        mainPanel.Controls.Add(buttonPanel, 0, 5);
        mainPanel.SetColumnSpan(buttonPanel, 2);

        Controls.Add(mainPanel);
        ResumeLayout(true);
    }

    private Panel CreateConnectionPanel(string labelText, out ComboBox comboBox)
    {
        var panel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 10, 5) };

        var label = new Label
        {
            Text = labelText,
            Dock = DockStyle.Top,
            Height = 20
        };

        comboBox = new ComboBox
        {
            Dock = DockStyle.Top,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        comboBox.SelectedIndexChanged += ConnectionCombo_SelectedIndexChanged;

        panel.Controls.Add(comboBox);
        panel.Controls.Add(label);

        return panel;
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        // Check if we received a TaskCreationContext with selected sites
        if (parameter is TaskCreationContext context)
        {
            _context = context;

            // Filter out subsites from selected sites
            _selectedSourceUrls = context.SelectedSites
                .Where(s => !IsSubsiteUrl(s.Url))
                .Select(s => s.Url)
                .ToList();

            var subsiteCount = context.SelectedSites.Count - _selectedSourceUrls.Count;

            if (_selectedSourceUrls.Count > 0)
            {
                var message = $"{_selectedSourceUrls.Count} root site(s) selected";
                if (subsiteCount > 0)
                {
                    message += $" ({subsiteCount} subsite(s) excluded)";
                }
                _selectedSitesLabel.Text = message;
                _selectedSitesLabel.ForeColor = SystemColors.ControlText;
            }
            else
            {
                _selectedSitesLabel.Text = "No root sites in selection (subsites are not supported)";
                _selectedSitesLabel.ForeColor = Color.DarkRed;
            }
        }

        await LoadConnectionsAsync();

        if (parameter is TenantPairTaskContext ctx && ctx.TenantPair.SitePairs.Count > 0)
        {
            _sitePairs.AddRange(ctx.TenantPair.SitePairs);
            RefreshSitePairsGrid();
            SelectConnectionById(_sourceConnectionCombo, ctx.TenantPair.SourceConnectionId);
            SelectConnectionById(_targetConnectionCombo, ctx.TenantPair.TargetConnectionId);
            SetStatus($"Loaded {ctx.TenantPair.SitePairs.Count} site pair(s) from tenant pair");
        }
    }

    private void SelectConnectionById(ComboBox combo, Guid connectionId)
    {
        var index = _connections.FindIndex(c => c.Id == connectionId);
        if (index >= 0)
            combo.SelectedIndex = index;
    }

    private async Task LoadConnectionsAsync()
    {
        ShowLoading("Loading connections...");

        try
        {
            _connections = await _connectionManager.GetAllConnectionsAsync();
            var tenantPairs = await _tenantPairService.GetAllPairsAsync();

            _sourceConnectionCombo.Items.Clear();
            _targetConnectionCombo.Items.Clear();

            foreach (var conn in _connections)
            {
                var displayName = $"{conn.Name} ({conn.TenantName})";
                _sourceConnectionCombo.Items.Add(displayName);
                _targetConnectionCombo.Items.Add(displayName);
            }

            if (_connections.Count > 0)
            {
                int sourceIndex = -1;
                int targetIndex = -1;

                // If we have a context, pre-select the source connection from context
                if (_context != null)
                {
                    sourceIndex = _connections.FindIndex(c => c.Id == _context.Connection.Id);
                }

                // Try to use the first tenant pair
                if (sourceIndex < 0 && tenantPairs.Count > 0)
                {
                    var firstPair = tenantPairs[0];
                    sourceIndex = _connections.FindIndex(c => c.Id == firstPair.SourceConnectionId);
                    targetIndex = _connections.FindIndex(c => c.Id == firstPair.TargetConnectionId);
                }

                // If no tenant pair, try to use connection roles
                if (sourceIndex < 0)
                {
                    sourceIndex = _connections.FindIndex(c => c.Role == TenantRole.Source);
                }
                if (targetIndex < 0)
                {
                    targetIndex = _connections.FindIndex(c => c.Role == TenantRole.Target);
                }

                // Fall back to first/second connection
                if (sourceIndex < 0) sourceIndex = 0;
                if (targetIndex < 0) targetIndex = _connections.Count > 1 ? (sourceIndex == 0 ? 1 : 0) : 0;

                _sourceConnectionCombo.SelectedIndex = sourceIndex;
                _targetConnectionCombo.SelectedIndex = targetIndex;
            }

            SetStatus($"Loaded {_connections.Count} connections");
        }
        finally
        {
            HideLoading();
        }
    }

    private void ConnectionCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        UpdateButtonStates();
    }

    private void UpdateButtonStates()
    {
        // Enable Generate Target URLs button if we have selected source URLs and both connections are selected
        _generateTargetUrlsButton.Enabled = _selectedSourceUrls.Count > 0 &&
                                            _sourceConnectionCombo.SelectedIndex >= 0 &&
                                            _targetConnectionCombo.SelectedIndex >= 0;

        _createTaskButton.Enabled = _sourceConnectionCombo.SelectedIndex >= 0 &&
                                    _targetConnectionCombo.SelectedIndex >= 0 &&
                                    _sitePairs.Count > 0 &&
                                    !string.IsNullOrWhiteSpace(_taskNameTextBox.Text);
        _clearPairsButton.Enabled = _sitePairs.Count > 0;
    }

    private void GenerateTargetUrlsButton_Click(object? sender, EventArgs e)
    {
        if (_sourceConnectionCombo.SelectedIndex < 0 || _targetConnectionCombo.SelectedIndex < 0)
        {
            MessageBox.Show(
                "Please select both source and target connections first.",
                "Validation Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        if (_selectedSourceUrls.Count == 0)
        {
            MessageBox.Show(
                "No source URLs available. Please select sites from the Site Collections screen.",
                "No Sites Selected",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        var sourceConnection = _connections[_sourceConnectionCombo.SelectedIndex];
        var targetConnection = _connections[_targetConnectionCombo.SelectedIndex];

        // Generate target URLs by replacing the source tenant domain with target tenant domain
        var generatedPairs = new List<SiteComparePair>();

        foreach (var sourceUrl in _selectedSourceUrls)
        {
            var targetUrl = sourceUrl.Replace(
                sourceConnection.TenantDomain,
                targetConnection.TenantDomain,
                StringComparison.OrdinalIgnoreCase);

            generatedPairs.Add(new SiteComparePair
            {
                SourceUrl = sourceUrl,
                TargetUrl = targetUrl
            });
        }

        // Clear existing pairs and add generated ones
        _sitePairs.Clear();
        _sitePairs.AddRange(generatedPairs);
        RefreshSitePairsGrid();

        SetStatus($"Generated {generatedPairs.Count} site pair(s) by mapping {sourceConnection.TenantName} to {targetConnection.TenantName}");
    }

    private void ImportCsvButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
            Title = "Select Site Pairs CSV File"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        try
        {
            var lines = File.ReadAllLines(dialog.FileName);
            var importedPairs = new List<SiteComparePair>();
            var errors = new List<string>();
            var lineNumber = 0;

            foreach (var line in lines)
            {
                lineNumber++;

                // Skip empty lines
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split(',');
                if (parts.Length < 2)
                {
                    // Try to detect if this is a header line
                    if (lineNumber == 1 && (line.Contains("Source", StringComparison.OrdinalIgnoreCase) ||
                                            line.Contains("URL", StringComparison.OrdinalIgnoreCase)))
                    {
                        continue;
                    }

                    errors.Add($"Line {lineNumber}: Not enough columns");
                    continue;
                }

                var sourceUrl = parts[0].Trim().Trim('"');
                var targetUrl = parts[1].Trim().Trim('"');

                // Validate URLs
                if (!Uri.TryCreate(sourceUrl, UriKind.Absolute, out var sourceUri))
                {
                    if (lineNumber == 1) continue; // Skip header row
                    errors.Add($"Line {lineNumber}: Invalid source URL '{sourceUrl}'");
                    continue;
                }

                if (!Uri.TryCreate(targetUrl, UriKind.Absolute, out var targetUri))
                {
                    if (lineNumber == 1) continue; // Skip header row
                    errors.Add($"Line {lineNumber}: Invalid target URL '{targetUrl}'");
                    continue;
                }

                // Check for subsites
                if (IsSubsite(sourceUri))
                {
                    errors.Add($"Line {lineNumber}: Source URL appears to be a subsite. Only root sites are supported.");
                    continue;
                }

                if (IsSubsite(targetUri))
                {
                    errors.Add($"Line {lineNumber}: Target URL appears to be a subsite. Only root sites are supported.");
                    continue;
                }

                importedPairs.Add(new SiteComparePair
                {
                    SourceUrl = sourceUrl,
                    TargetUrl = targetUrl
                });
            }

            if (errors.Count > 0 && importedPairs.Count == 0)
            {
                MessageBox.Show(
                    $"Failed to import CSV file:\n\n{string.Join("\n", errors.Take(10))}",
                    "Import Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            // Add imported pairs
            _sitePairs.AddRange(importedPairs);
            RefreshSitePairsGrid();

            var message = $"Imported {importedPairs.Count} site pair(s).";
            if (errors.Count > 0)
            {
                message += $"\n\n{errors.Count} line(s) had errors and were skipped.";
            }

            SetStatus(message);

            if (errors.Count > 0)
            {
                MessageBox.Show(
                    $"{message}\n\nErrors:\n{string.Join("\n", errors.Take(5))}",
                    "Import Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to read CSV file: {ex.Message}",
                "Import Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static bool IsSubsiteUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return false;
        return IsSubsite(uri);
    }

    private static bool IsSubsite(Uri uri)
    {
        var pathParts = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);

        // Check if it's a /sites/ or /teams/ path with more than 2 segments
        if (pathParts.Length > 2)
        {
            if (pathParts[0].Equals("sites", StringComparison.OrdinalIgnoreCase) ||
                pathParts[0].Equals("teams", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private void ClearPairsButton_Click(object? sender, EventArgs e)
    {
        var result = MessageBox.Show(
            "Are you sure you want to clear all site pairs?",
            "Clear Site Pairs",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            _sitePairs.Clear();
            RefreshSitePairsGrid();
            SetStatus("Site pairs cleared");
        }
    }

    private void RefreshSitePairsGrid()
    {
        _sitePairsGrid.Rows.Clear();

        foreach (var pair in _sitePairs)
        {
            _sitePairsGrid.Rows.Add(pair.SourceUrl, pair.TargetUrl);
        }

        UpdateButtonStates();
    }

    private async void CreateTaskButton_Click(object? sender, EventArgs e)
    {
        if (_sourceConnectionCombo.SelectedIndex < 0 ||
            _targetConnectionCombo.SelectedIndex < 0 ||
            _sitePairs.Count == 0)
        {
            MessageBox.Show(
                "Please select source and target connections and add at least one site pair.",
                "Validation Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(_taskNameTextBox.Text))
        {
            MessageBox.Show(
                "Please enter a task name.",
                "Validation Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        var sourceConnection = _connections[_sourceConnectionCombo.SelectedIndex];
        var targetConnection = _connections[_targetConnectionCombo.SelectedIndex];

        // Check authentication for both connections
        if (!await EnsureAuthenticationAsync(sourceConnection, "source"))
            return;

        if (!await EnsureAuthenticationAsync(targetConnection, "target"))
            return;

        // Build configuration
        var config = new NavigationSettingsConfiguration
        {
            SourceConnectionId = sourceConnection.Id,
            TargetConnectionId = targetConnection.Id,
            SitePairs = _sitePairs.ToList()
        };

        // Create task
        var jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Converters = { new JsonStringEnumConverter() }
        };

        var task = new TaskDefinition
        {
            Name = _taskNameTextBox.Text.Trim(),
            Type = TaskType.NavigationSettingsSync,
            ConnectionId = sourceConnection.Id,
            ConfigurationJson = JsonSerializer.Serialize(config, jsonOptions),
            Status = Models.TaskStatus.Pending
        };

        await _taskService.SaveTaskAsync(task);

        SetStatus($"Task '{task.Name}' created with {_sitePairs.Count} site pairs");

        // Navigate to detail screen
        await NavigationService!.NavigateToAsync<NavigationSettingsDetailScreen>(task);
    }

    private async Task<bool> EnsureAuthenticationAsync(Connection connection, string connectionLabel)
    {
        if (_authService.HasStoredCredentials(connection.CookieDomain))
            return true;

        var result = MessageBox.Show(
            $"Authentication required for {connectionLabel} connection '{connection.Name}'.\n\nWould you like to sign in?",
            "Authentication Required",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result != DialogResult.Yes)
            return false;

        SetStatus($"Opening sign-in window for {connection.Name}...");

        using var loginForm = new LoginForm(connection.PrimaryUrl);
        if (loginForm.ShowDialog(FindForm()) != DialogResult.OK || loginForm.CapturedCookies == null)
        {
            SetStatus("Authentication cancelled");
            return false;
        }

        _authService.StoreCookies(loginForm.CapturedCookies);
        SetStatus($"Authenticated to {connection.CookieDomain}");
        return true;
    }
}

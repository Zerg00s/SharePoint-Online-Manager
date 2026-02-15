using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for configuring a new List Compare task.
/// </summary>
public class ListCompareConfigScreen : BaseScreen
{
    private ComboBox _sourceConnectionCombo = null!;
    private ComboBox _targetConnectionCombo = null!;
    private Button _importCsvButton = null!;
    private DataGridView _sitePairsGrid = null!;
    private CheckedListBox _exclusionsListBox = null!;
    private CheckBox _includeSiteAssetsCheckBox = null!;
    private CheckBox _includeHiddenListsCheckBox = null!;
    private ComboBox _thresholdTypeCombo = null!;
    private NumericUpDown _thresholdValueNumeric = null!;
    private TextBox _taskNameTextBox = null!;
    private Button _createTaskButton = null!;
    private Button _clearPairsButton = null!;

    private IConnectionManager _connectionManager = null!;
    private IAuthenticationService _authService = null!;
    private ITaskService _taskService = null!;
    private ITenantPairService _tenantPairService = null!;
    private List<Connection> _connections = [];
    private List<SiteComparePair> _sitePairs = [];

    public override string ScreenTitle => "Create List Compare Task";

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
            RowCount = 7,
            Padding = new Padding(10)
        };
        mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        mainPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 70)); // Connections row
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45)); // Import button row
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 40)); // Site pairs grid
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30)); // Exclusions
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80)); // Threshold settings
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60)); // Task name
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45)); // Create button

        // Row 0: Connection selectors
        var sourceConnPanel = CreateConnectionPanel("Source Connection:", out _sourceConnectionCombo);
        var targetConnPanel = CreateConnectionPanel("Target Connection:", out _targetConnectionCombo);
        mainPanel.Controls.Add(sourceConnPanel, 0, 0);
        mainPanel.Controls.Add(targetConnPanel, 1, 0);

        // Row 1: Import CSV button
        var importPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

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

        importPanel.Controls.AddRange(new Control[] { _importCsvButton, _clearPairsButton, csvInfoLabel });
        mainPanel.Controls.Add(importPanel, 0, 1);
        mainPanel.SetColumnSpan(importPanel, 2);

        // Row 2: Site pairs grid
        var sitePairsPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
        var sitePairsLabel = new Label
        {
            Text = "Site Pairs:",
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
        mainPanel.Controls.Add(sitePairsPanel, 0, 2);
        mainPanel.SetColumnSpan(sitePairsPanel, 2);

        // Row 3: Exclusions panel
        var exclusionsPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
        var exclusionsLabel = new Label
        {
            Text = "Excluded System Lists (always excluded):",
            Dock = DockStyle.Top,
            Height = 20
        };

        _exclusionsListBox = new CheckedListBox
        {
            Dock = DockStyle.Fill,
            CheckOnClick = true
        };

        // Add default exclusions (all checked by default)
        foreach (var list in ListCompareConfiguration.DefaultExcludedLists)
        {
            _exclusionsListBox.Items.Add(list, true);
        }

        exclusionsPanel.Controls.Add(_exclusionsListBox);
        exclusionsPanel.Controls.Add(exclusionsLabel);
        mainPanel.Controls.Add(exclusionsPanel, 0, 3);
        mainPanel.SetColumnSpan(exclusionsPanel, 2);

        // Row 4: Options panel (threshold + checkboxes)
        var optionsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        _includeSiteAssetsCheckBox = new CheckBox
        {
            Text = "Include Site Assets",
            AutoSize = true,
            Checked = false,
            Margin = new Padding(0, 5, 20, 0)
        };

        _includeHiddenListsCheckBox = new CheckBox
        {
            Text = "Include Hidden Lists",
            AutoSize = true,
            Checked = false,
            Margin = new Padding(0, 5, 40, 0)
        };

        var thresholdLabel = new Label
        {
            Text = "Threshold:",
            AutoSize = true,
            Padding = new Padding(0, 8, 5, 0)
        };

        _thresholdTypeCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 120,
            Margin = new Padding(0, 5, 5, 0)
        };
        _thresholdTypeCombo.Items.AddRange(new object[] { "Percentage", "Absolute Count" });
        _thresholdTypeCombo.SelectedIndex = 0;

        _thresholdValueNumeric = new NumericUpDown
        {
            Minimum = 0,
            Maximum = 100000,
            Value = 10,
            Width = 80,
            Margin = new Padding(0, 5, 0, 0)
        };

        var thresholdSuffixLabel = new Label
        {
            Text = "%",
            AutoSize = true,
            Padding = new Padding(0, 8, 0, 0),
            Name = "ThresholdSuffix"
        };

        _thresholdTypeCombo.SelectedIndexChanged += (s, e) =>
        {
            if (_thresholdTypeCombo.SelectedIndex == 0)
            {
                thresholdSuffixLabel.Text = "%";
                _thresholdValueNumeric.Value = Math.Min(_thresholdValueNumeric.Value, 100);
                _thresholdValueNumeric.Maximum = 100;
            }
            else
            {
                thresholdSuffixLabel.Text = "items";
                _thresholdValueNumeric.Maximum = 100000;
                _thresholdValueNumeric.Value = 100;
            }
        };

        optionsPanel.Controls.AddRange(new Control[]
        {
            _includeSiteAssetsCheckBox,
            _includeHiddenListsCheckBox,
            thresholdLabel,
            _thresholdTypeCombo,
            _thresholdValueNumeric,
            thresholdSuffixLabel
        });
        mainPanel.Controls.Add(optionsPanel, 0, 4);
        mainPanel.SetColumnSpan(optionsPanel, 2);

        // Row 5: Task name
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
            Text = $"List Compare - {DateTime.Now:yyyy-MM-dd HH:mm}"
        };

        taskNamePanel.Controls.Add(_taskNameTextBox);
        taskNamePanel.Controls.Add(taskNameLabel);
        mainPanel.Controls.Add(taskNamePanel, 0, 5);
        mainPanel.SetColumnSpan(taskNamePanel, 2);

        // Row 6: Create button
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
        mainPanel.Controls.Add(buttonPanel, 0, 6);
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
        await LoadConnectionsAsync();
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

                // First, try to use the first tenant pair
                if (tenantPairs.Count > 0)
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
                if (targetIndex < 0) targetIndex = _connections.Count > 1 ? 1 : 0;

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
        UpdateCreateButtonState();
    }

    private void UpdateCreateButtonState()
    {
        _createTaskButton.Enabled = _sourceConnectionCombo.SelectedIndex >= 0 &&
                                    _targetConnectionCombo.SelectedIndex >= 0 &&
                                    _sitePairs.Count > 0 &&
                                    !string.IsNullOrWhiteSpace(_taskNameTextBox.Text);
        _clearPairsButton.Enabled = _sitePairs.Count > 0;
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

                // Skip empty lines and header
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split(',');
                if (parts.Length < 2)
                {
                    // Try to detect if this is a header line
                    if (lineNumber == 1 && (line.Contains("Source", StringComparison.OrdinalIgnoreCase) ||
                                            line.Contains("URL", StringComparison.OrdinalIgnoreCase)))
                    {
                        continue; // Skip header
                    }

                    errors.Add($"Line {lineNumber}: Not enough columns");
                    continue;
                }

                var sourceUrl = parts[0].Trim().Trim('"');
                var targetUrl = parts[1].Trim().Trim('"');

                // Validate URLs
                if (!Uri.TryCreate(sourceUrl, UriKind.Absolute, out _))
                {
                    if (lineNumber == 1) continue; // Skip header row
                    errors.Add($"Line {lineNumber}: Invalid source URL '{sourceUrl}'");
                    continue;
                }

                if (!Uri.TryCreate(targetUrl, UriKind.Absolute, out _))
                {
                    if (lineNumber == 1) continue; // Skip header row
                    errors.Add($"Line {lineNumber}: Invalid target URL '{targetUrl}'");
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

        UpdateCreateButtonState();
    }

    private async void CreateTaskButton_Click(object? sender, EventArgs e)
    {
        if (_sourceConnectionCombo.SelectedIndex < 0 ||
            _targetConnectionCombo.SelectedIndex < 0 ||
            _sitePairs.Count == 0)
        {
            MessageBox.Show(
                "Please select source and target connections and import at least one site pair.",
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
        var config = new ListCompareConfiguration
        {
            SourceConnectionId = sourceConnection.Id,
            TargetConnectionId = targetConnection.Id,
            SitePairs = _sitePairs.ToList(),
            ExcludedLists = GetExcludedLists(),
            IncludeSiteAssets = _includeSiteAssetsCheckBox.Checked,
            IncludeHiddenLists = _includeHiddenListsCheckBox.Checked,
            ThresholdType = _thresholdTypeCombo.SelectedIndex == 0
                ? ThresholdType.Percentage
                : ThresholdType.AbsoluteCount,
            ThresholdValue = (int)_thresholdValueNumeric.Value
        };

        // Create task - use camelCase to match TaskService deserialization
        var jsonOptions = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Converters = { new JsonStringEnumConverter() }
        };

        var task = new TaskDefinition
        {
            Name = _taskNameTextBox.Text.Trim(),
            Type = TaskType.ListCompare,
            ConnectionId = sourceConnection.Id, // Primary connection for display
            ConfigurationJson = JsonSerializer.Serialize(config, jsonOptions),
            Status = Models.TaskStatus.Pending
        };

        await _taskService.SaveTaskAsync(task);

        SetStatus($"Task '{task.Name}' created with {_sitePairs.Count} site pairs");

        // Navigate to detail screen
        await NavigationService!.NavigateToAsync<ListCompareDetailScreen>(task);
    }

    private List<string> GetExcludedLists()
    {
        var excluded = new List<string>();

        for (int i = 0; i < _exclusionsListBox.Items.Count; i++)
        {
            if (_exclusionsListBox.GetItemChecked(i))
            {
                excluded.Add(_exclusionsListBox.Items[i].ToString()!);
            }
        }

        // Add optional exclusions if not including Site Assets
        if (!_includeSiteAssetsCheckBox.Checked)
        {
            excluded.AddRange(ListCompareConfiguration.OptionalExcludedLists);
        }

        return excluded;
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

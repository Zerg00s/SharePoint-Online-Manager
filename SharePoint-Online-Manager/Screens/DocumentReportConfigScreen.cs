using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for configuring a new Document Report task.
/// </summary>
public class DocumentReportConfigScreen : BaseScreen
{
    private ComboBox _connectionCombo = null!;
    private Button _importCsvButton = null!;
    private DataGridView _sitesGrid = null!;
    private CheckBox _includeHiddenCheckBox = null!;
    private CheckBox _includeSubfoldersCheckBox = null!;
    private CheckBox _includeVersionsCheckBox = null!;
    private TextBox _extensionFilterTextBox = null!;
    private TextBox _taskNameTextBox = null!;
    private Button _createTaskButton = null!;
    private Button _clearSitesButton = null!;

    private IConnectionManager _connectionManager = null!;
    private IAuthenticationService _authService = null!;
    private ITaskService _taskService = null!;
    private List<Connection> _connections = [];
    private List<string> _siteUrls = [];

    public override string ScreenTitle => "Create Document Report Task";

    protected override void OnInitialize()
    {
        _connectionManager = GetRequiredService<IConnectionManager>();
        _authService = GetRequiredService<IAuthenticationService>();
        _taskService = GetRequiredService<ITaskService>();
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
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));  // Connection row
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));  // Import button row
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50));   // Sites grid
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 100)); // Options panel
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));  // Extension filter
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));  // Task name
        mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));  // Create button

        // Row 0: Connection selector
        var connectionPanel = CreateConnectionPanel("Connection:", out _connectionCombo);
        mainPanel.Controls.Add(connectionPanel, 0, 0);
        mainPanel.SetColumnSpan(connectionPanel, 2);

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

        _clearSitesButton = new Button
        {
            Text = "Clear All",
            Size = new Size(80, 28),
            Enabled = false
        };
        _clearSitesButton.Click += ClearSitesButton_Click;

        var csvInfoLabel = new Label
        {
            Text = "CSV format: Site URL (one per line or comma-separated)",
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
            Padding = new Padding(10, 8, 0, 0)
        };

        importPanel.Controls.AddRange(new Control[] { _importCsvButton, _clearSitesButton, csvInfoLabel });
        mainPanel.Controls.Add(importPanel, 0, 1);
        mainPanel.SetColumnSpan(importPanel, 2);

        // Row 2: Sites grid
        var sitesPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
        var sitesLabel = new Label
        {
            Text = "Target Sites:",
            Dock = DockStyle.Top,
            Height = 20
        };

        _sitesGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = true,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };
        _sitesGrid.Columns.Add("SiteUrl", "Site URL");

        sitesPanel.Controls.Add(_sitesGrid);
        sitesPanel.Controls.Add(sitesLabel);
        mainPanel.Controls.Add(sitesPanel, 0, 2);
        mainPanel.SetColumnSpan(sitesPanel, 2);

        // Row 3: Options panel
        var optionsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        _includeHiddenCheckBox = new CheckBox
        {
            Text = "Include Hidden Libraries",
            AutoSize = true,
            Checked = false,
            Margin = new Padding(0, 5, 30, 0)
        };

        _includeSubfoldersCheckBox = new CheckBox
        {
            Text = "Include Subfolders",
            AutoSize = true,
            Checked = true,
            Margin = new Padding(0, 5, 30, 0)
        };

        _includeVersionsCheckBox = new CheckBox
        {
            Text = "Include Version Count",
            AutoSize = true,
            Checked = true,
            Margin = new Padding(0, 5, 0, 0)
        };

        optionsPanel.Controls.AddRange(new Control[]
        {
            _includeHiddenCheckBox,
            _includeSubfoldersCheckBox,
            _includeVersionsCheckBox
        });
        mainPanel.Controls.Add(optionsPanel, 0, 3);
        mainPanel.SetColumnSpan(optionsPanel, 2);

        // Row 4: Extension filter
        var extensionPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 5) };
        var extensionLabel = new Label
        {
            Text = "Extension Filter (comma-separated, e.g., pdf,docx,xlsx - leave empty for all):",
            Dock = DockStyle.Top,
            Height = 20
        };

        _extensionFilterTextBox = new TextBox
        {
            Dock = DockStyle.Top,
            PlaceholderText = "e.g., pdf, docx, xlsx (leave empty for all files)"
        };

        extensionPanel.Controls.Add(_extensionFilterTextBox);
        extensionPanel.Controls.Add(extensionLabel);
        mainPanel.Controls.Add(extensionPanel, 0, 4);
        mainPanel.SetColumnSpan(extensionPanel, 2);

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
            Text = $"Document Report - {DateTime.Now:yyyy-MM-dd HH:mm}"
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

            _connectionCombo.Items.Clear();

            foreach (var conn in _connections)
            {
                var displayName = $"{conn.Name} ({conn.TenantName})";
                _connectionCombo.Items.Add(displayName);
            }

            if (_connections.Count > 0)
            {
                _connectionCombo.SelectedIndex = 0;
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
        _createTaskButton.Enabled = _connectionCombo.SelectedIndex >= 0 &&
                                    _siteUrls.Count > 0 &&
                                    !string.IsNullOrWhiteSpace(_taskNameTextBox.Text);
        _clearSitesButton.Enabled = _siteUrls.Count > 0;
    }

    private void ImportCsvButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
            Title = "Select Sites File"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        try
        {
            var lines = File.ReadAllLines(dialog.FileName);
            var importedSites = new List<string>();
            var errors = new List<string>();
            var lineNumber = 0;

            foreach (var line in lines)
            {
                lineNumber++;

                // Skip empty lines
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Split by comma in case multiple URLs per line
                var urls = line.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var urlPart in urls)
                {
                    var url = urlPart.Trim().Trim('"');

                    // Skip header row
                    if (url.Equals("Site URL", StringComparison.OrdinalIgnoreCase) ||
                        url.Equals("SiteUrl", StringComparison.OrdinalIgnoreCase) ||
                        url.Equals("URL", StringComparison.OrdinalIgnoreCase) ||
                        url.Equals("Site", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // Validate URL
                    if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                        (uri.Scheme != "http" && uri.Scheme != "https"))
                    {
                        errors.Add($"Line {lineNumber}: Invalid URL '{url}'");
                        continue;
                    }

                    // Avoid duplicates
                    if (!importedSites.Contains(url, StringComparer.OrdinalIgnoreCase) &&
                        !_siteUrls.Contains(url, StringComparer.OrdinalIgnoreCase))
                    {
                        importedSites.Add(url);
                    }
                }
            }

            if (errors.Count > 0 && importedSites.Count == 0)
            {
                MessageBox.Show(
                    $"Failed to import file:\n\n{string.Join("\n", errors.Take(10))}",
                    "Import Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            // Add imported sites
            _siteUrls.AddRange(importedSites);
            RefreshSitesGrid();

            var message = $"Imported {importedSites.Count} site(s).";
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
                $"Failed to read file: {ex.Message}",
                "Import Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void ClearSitesButton_Click(object? sender, EventArgs e)
    {
        var result = MessageBox.Show(
            "Are you sure you want to clear all sites?",
            "Clear Sites",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            _siteUrls.Clear();
            RefreshSitesGrid();
            SetStatus("Sites cleared");
        }
    }

    private void RefreshSitesGrid()
    {
        _sitesGrid.Rows.Clear();

        foreach (var url in _siteUrls)
        {
            _sitesGrid.Rows.Add(url);
        }

        UpdateCreateButtonState();
    }

    private async void CreateTaskButton_Click(object? sender, EventArgs e)
    {
        if (_connectionCombo.SelectedIndex < 0 || _siteUrls.Count == 0)
        {
            MessageBox.Show(
                "Please select a connection and import at least one site.",
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

        var connection = _connections[_connectionCombo.SelectedIndex];

        // Check authentication
        if (!await EnsureAuthenticationAsync(connection))
            return;

        // Build configuration
        var config = new DocumentReportConfiguration
        {
            ConnectionId = connection.Id,
            TargetSiteUrls = _siteUrls.ToList(),
            IncludeHiddenLibraries = _includeHiddenCheckBox.Checked,
            IncludeSubfolders = _includeSubfoldersCheckBox.Checked,
            IncludeVersionCount = _includeVersionsCheckBox.Checked,
            ExtensionFilter = _extensionFilterTextBox.Text.Trim()
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
            Type = TaskType.DocumentReport,
            ConnectionId = connection.Id,
            TargetSiteUrls = _siteUrls.ToList(),
            ConfigurationJson = JsonSerializer.Serialize(config, jsonOptions),
            Status = Models.TaskStatus.Pending
        };

        await _taskService.SaveTaskAsync(task);

        SetStatus($"Task '{task.Name}' created with {_siteUrls.Count} sites");

        // Navigate to detail screen
        await NavigationService!.NavigateToAsync<DocumentReportDetailScreen>(task);
    }

    private async Task<bool> EnsureAuthenticationAsync(Connection connection)
    {
        if (_authService.HasStoredCredentials(connection.CookieDomain))
            return true;

        var result = MessageBox.Show(
            $"Authentication required for connection '{connection.Name}'.\n\nWould you like to sign in?",
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

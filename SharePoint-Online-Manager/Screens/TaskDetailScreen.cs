using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing task details and results.
/// </summary>
public class TaskDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _exportButton = null!;
    private Button _deleteButton = null!;
    private Button _manageExclusionsButton = null!;
    private DataGridView _resultsGrid = null!;
    private TextBox _logTextBox = null!;
    private TextBox _filterTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private Label _filterLabel = null!;

    private TaskDefinition _task = null!;
    private TaskResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private HashSet<string> _excludedLists = new(StringComparer.OrdinalIgnoreCase);

    public override string ScreenTitle => _task?.Name ?? "Task Details";

    protected override void OnInitialize()
    {
        _taskService = GetRequiredService<ITaskService>();
        _authService = GetRequiredService<IAuthenticationService>();
        _connectionManager = GetRequiredService<IConnectionManager>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();

        // Header panel
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 100
        };

        _taskNameLabel = new Label
        {
            Location = new Point(0, 5),
            AutoSize = true,
            Font = new Font(Font.FontFamily, 14F, FontStyle.Bold)
        };

        _taskInfoLabel = new Label
        {
            Location = new Point(0, 35),
            AutoSize = true
        };

        var buttonPanel = new FlowLayoutPanel
        {
            Location = new Point(0, 65),
            Size = new Size(700, 35),
            FlowDirection = FlowDirection.LeftToRight
        };

        _runButton = new Button
        {
            Text = "Run Task",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0)
        };
        _runButton.Click += RunButton_Click;

        _exportButton = new Button
        {
            Text = "Export to CSV",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false
        };
        _exportButton.Click += ExportButton_Click;

        _deleteButton = new Button
        {
            Text = "Delete Task",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0)
        };
        _deleteButton.Click += DeleteButton_Click;

        _manageExclusionsButton = new Button
        {
            Text = "Manage Exclusions",
            Size = new Size(130, 28)
        };
        _manageExclusionsButton.Click += ManageExclusionsButton_Click;

        buttonPanel.Controls.AddRange(new Control[] { _runButton, _exportButton, _deleteButton, _manageExclusionsButton });

        headerPanel.Controls.AddRange(new Control[] { _taskNameLabel, _taskInfoLabel, buttonPanel });

        // Progress panel
        var progressPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            Visible = false
        };

        _progressBar = new ProgressBar
        {
            Location = new Point(0, 5),
            Size = new Size(600, 20)
        };

        _progressLabel = new Label
        {
            Location = new Point(0, 28),
            AutoSize = true,
            Text = "Ready"
        };

        progressPanel.Controls.AddRange(new Control[] { _progressBar, _progressLabel });

        // Tab control for results and log
        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        // Results tab
        var resultsTab = new TabPage("Results");

        // Filter panel at top of results tab
        var filterPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 35,
            Padding = new Padding(5)
        };

        _filterLabel = new Label
        {
            Text = "Filter:",
            AutoSize = true,
            Location = new Point(5, 10)
        };

        _filterTextBox = new TextBox
        {
            Location = new Point(50, 7),
            Size = new Size(300, 23),
            PlaceholderText = "Type to filter by list name..."
        };
        _filterTextBox.TextChanged += FilterTextBox_TextChanged;

        filterPanel.Controls.Add(_filterLabel);
        filterPanel.Controls.Add(_filterTextBox);

        _resultsGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        };

        _resultsGrid.Columns.Add("SiteUrl", "Site URL");
        _resultsGrid.Columns.Add("SiteTitle", "Site Title");
        _resultsGrid.Columns.Add("ListTitle", "List Title");
        _resultsGrid.Columns.Add("ListType", "List Type");
        _resultsGrid.Columns.Add("ItemCount", "Item Count");
        _resultsGrid.Columns.Add("Hidden", "Hidden");
        _resultsGrid.Columns.Add("LastModified", "Last Modified");

        resultsTab.Controls.Add(_resultsGrid);
        resultsTab.Controls.Add(filterPanel);

        // Log tab
        var logTab = new TabPage("Execution Log");
        _logTextBox = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Both,
            ReadOnly = true,
            Font = new Font("Consolas", 9F),
            WordWrap = false
        };
        logTab.Controls.Add(_logTextBox);

        _tabControl.TabPages.AddRange([resultsTab, logTab]);

        Controls.Add(_tabControl);
        Controls.Add(progressPanel);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        TaskDefinition? task = null;
        var executeImmediately = false;

        if (parameter is TaskExecutionParameter execParam)
        {
            task = execParam.Task;
            executeImmediately = execParam.ExecuteImmediately;
        }
        else if (parameter is TaskDefinition taskDef)
        {
            task = taskDef;
        }

        if (task == null)
        {
            MessageBox.Show("No task specified.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            await NavigationService!.GoBackAsync();
            return;
        }

        _task = task;
        await RefreshTaskDetailsAsync();

        if (executeImmediately)
        {
            await ExecuteTaskAsync();
        }
    }

    private async Task RefreshTaskDetailsAsync()
    {
        _taskNameLabel.Text = _task.Name;

        var connection = await _connectionManager.GetConnectionAsync(_task.ConnectionId);
        var connectionName = connection?.Name ?? "Unknown";

        _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Sites: {_task.TotalSites} | " +
                              $"Status: {_task.StatusDescription} | Connection: {connectionName}";

        // Load latest result
        _currentResult = await _taskService.GetLatestTaskResultAsync(_task.Id);

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportButton.Enabled = true;
        }
        else
        {
            _resultsGrid.Rows.Clear();
            _logTextBox.Text = "No results yet. Click 'Run Task' to execute.";
            _exportButton.Enabled = false;
        }
    }

    private void DisplayResults(TaskResult result)
    {
        DisplayResults(result, _filterTextBox?.Text ?? string.Empty);
    }

    private void DisplayResults(TaskResult result, string filter)
    {
        _resultsGrid.Rows.Clear();

        var items = result.GetAllListItems()
            .Where(item => !_excludedLists.Contains(item.ListTitle));

        // Apply text filter if specified
        if (!string.IsNullOrWhiteSpace(filter))
        {
            items = items.Where(item =>
                item.ListTitle.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                item.SiteTitle.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                item.SiteUrl.Contains(filter, StringComparison.OrdinalIgnoreCase));
        }

        foreach (var item in items)
        {
            var rowIndex = _resultsGrid.Rows.Add(
                item.SiteUrl,
                item.SiteTitle,
                item.ListTitle,
                item.ListType,
                item.ItemCount,
                item.Hidden ? "Yes" : "No",
                item.LastModified.ToString("g")
            );

            if (item.Hidden)
            {
                _resultsGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.Gray;
            }
        }

        // Update status with count
        var totalCount = result.GetAllListItems().Count();
        var excludedCount = result.GetAllListItems().Count(i => _excludedLists.Contains(i.ListTitle));
        var displayedCount = _resultsGrid.Rows.Count;

        if (_excludedLists.Count > 0 || !string.IsNullOrWhiteSpace(filter))
        {
            SetStatus($"Showing {displayedCount} of {totalCount} lists ({excludedCount} excluded, {totalCount - displayedCount - excludedCount} filtered)");
        }

        // Display log
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);
    }

    private void FilterTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayResults(_currentResult, _filterTextBox.Text);
        }
    }

    private async void RunButton_Click(object? sender, EventArgs e)
    {
        if (_runButton.Text == "Cancel")
        {
            _cancellationTokenSource?.Cancel();
            return;
        }

        await ExecuteTaskAsync();
    }

    private async Task ExecuteTaskAsync()
    {
        var connection = await _connectionManager.GetConnectionAsync(_task.ConnectionId);
        if (connection == null)
        {
            MessageBox.Show("Connection not found.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Check authentication
        if (!_authService.HasStoredCredentials(connection.CookieDomain))
        {
            var result = MessageBox.Show(
                "Authentication required. Would you like to sign in?",
                "Authentication Required",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            using var loginForm = new LoginForm(connection.PrimaryUrl);
            if (loginForm.ShowDialog(FindForm()) != DialogResult.OK || loginForm.CapturedCookies == null)
            {
                return;
            }

            _authService.StoreCookies(loginForm.CapturedCookies);
        }

        // Setup for execution
        _runButton.Text = "Cancel";
        _deleteButton.Enabled = false;
        _exportButton.Enabled = false;

        var progressPanel = Controls.OfType<Panel>().Skip(1).First();
        progressPanel.Visible = true;
        _progressBar.Value = 0;
        _progressLabel.Text = "Starting...";

        _cancellationTokenSource = new CancellationTokenSource();
        _resultsGrid.Rows.Clear();
        _logTextBox.Clear();

        var progress = new Progress<TaskProgress>(p =>
        {
            _progressBar.Value = p.PercentComplete;
            _progressLabel.Text = p.Message;
        });

        try
        {
            _currentResult = await _taskService.ExecuteTaskAsync(
                _task,
                _authService,
                progress,
                _cancellationTokenSource.Token);

            DisplayResults(_currentResult);

            if (_currentResult.Success)
            {
                SetStatus($"Task completed successfully. Processed {_currentResult.TotalSitesProcessed} sites.");
            }
            else
            {
                SetStatus($"Task completed with errors. {_currentResult.FailedSites} site(s) failed.");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Task execution failed: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Task execution failed");
        }
        finally
        {
            _runButton.Text = "Run Task";
            _deleteButton.Enabled = true;
            _exportButton.Enabled = _currentResult != null;
            progressPanel.Visible = false;
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;

            // Refresh task details to show updated status
            var updatedTask = await _taskService.GetTaskAsync(_task.Id);
            if (updatedTask != null)
            {
                _task = updatedTask;
                _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Sites: {_task.TotalSites} | " +
                                      $"Status: {_task.StatusDescription}";
            }
        }
    }

    private void ExportButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{_task.Name}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var exporter = GetRequiredService<CsvExporter>();
                exporter.ExportListReport(_currentResult, dialog.FileName, _excludedLists);
                SetStatus($"Exported to {dialog.FileName}" + (_excludedLists.Count > 0 ? $" ({_excludedLists.Count} lists excluded)" : ""));

                var result = MessageBox.Show(
                    "Export completed. Would you like to open the file?",
                    "Export Complete",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = dialog.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private async void DeleteButton_Click(object? sender, EventArgs e)
    {
        var result = MessageBox.Show(
            $"Are you sure you want to delete the task '{_task.Name}'?\n\nThis will also delete all saved results.",
            "Delete Task",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            await _taskService.DeleteTaskAsync(_task.Id);
            SetStatus($"Task '{_task.Name}' deleted");
            await NavigationService!.GoBackAsync();
        }
    }

    private void ManageExclusionsButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
        {
            MessageBox.Show("Run the task first to see the list of available lists to exclude.",
                "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Get all unique list titles from the results
        var allLists = _currentResult.GetAllListItems()
            .Select(i => i.ListTitle)
            .Distinct()
            .OrderBy(t => t)
            .ToList();

        using var dialog = new ListExclusionDialog(allLists, _excludedLists);
        if (dialog.ShowDialog(FindForm()) == DialogResult.OK)
        {
            _excludedLists = new HashSet<string>(dialog.ExcludedLists, StringComparer.OrdinalIgnoreCase);
            DisplayResults(_currentResult);
            SetStatus($"Exclusions updated: {_excludedLists.Count} lists excluded");
        }
    }

    public override Task<bool> OnNavigatingFromAsync()
    {
        if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
        {
            var result = MessageBox.Show(
                "A task is currently running. Are you sure you want to leave?",
                "Task Running",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
            {
                return Task.FromResult(false);
            }

            _cancellationTokenSource.Cancel();
        }

        return Task.FromResult(true);
    }
}

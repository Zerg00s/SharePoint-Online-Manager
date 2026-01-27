using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;
using System.Text.Json;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for executing and viewing Set Site State task results.
/// </summary>
public class SetSiteStateDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _deleteButton = null!;
    private DataGridView _resultsGrid = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private Panel _progressPanel = null!;

    private TaskDefinition _task = null!;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;

    public override string ScreenTitle => _task?.Name ?? "Set Site State";

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
            Text = "\u25B6 Run Task",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _runButton.FlatAppearance.BorderSize = 0;
        _runButton.Click += RunButton_Click;

        _deleteButton = new Button
        {
            Text = "\U0001F5D1 Delete Task",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            ForeColor = Color.DarkRed
        };
        _deleteButton.FlatAppearance.BorderColor = Color.DarkRed;
        _deleteButton.Click += DeleteButton_Click;

        buttonPanel.Controls.AddRange([_runButton, _deleteButton]);
        headerPanel.Controls.AddRange([_taskNameLabel, _taskInfoLabel, buttonPanel]);

        // Progress panel
        _progressPanel = new Panel
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

        _progressPanel.Controls.AddRange([_progressBar, _progressLabel]);

        // Results grid
        _resultsGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        _resultsGrid.Columns.Add("SiteUrl", "Site URL");
        _resultsGrid.Columns.Add("SiteTitle", "Title");
        _resultsGrid.Columns.Add("PreviousState", "Previous State");
        _resultsGrid.Columns.Add("NewState", "New State");
        _resultsGrid.Columns.Add("Status", "Status");
        _resultsGrid.Columns.Add("Error", "Error");

        Controls.Add(_resultsGrid);
        Controls.Add(_progressPanel);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        if (parameter is TaskDefinition task)
        {
            _task = task;
        }
        else if (parameter is Guid taskId)
        {
            _task = await _taskService.GetTaskAsync(taskId)
                ?? throw new InvalidOperationException("Task not found");
        }
        else
        {
            MessageBox.Show("No task specified.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            await NavigationService!.GoBackAsync();
            return;
        }

        UpdateUI();
    }

    private void UpdateUI()
    {
        _taskNameLabel.Text = _task.Name;

        var targetState = GetTargetState();
        _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Sites: {_task.TargetSiteUrls.Count} | " +
                              $"Target State: {targetState} | Status: {_task.Status}";

        UpdateTitle();

        // Pre-populate the grid with target sites
        _resultsGrid.Rows.Clear();
        foreach (var siteUrl in _task.TargetSiteUrls)
        {
            _resultsGrid.Rows.Add(siteUrl, "", "", targetState, "Pending", "");
        }
    }

    private string GetTargetState()
    {
        if (string.IsNullOrEmpty(_task.ConfigurationJson))
            return "Unknown";

        try
        {
            using var doc = JsonDocument.Parse(_task.ConfigurationJson);
            if (doc.RootElement.TryGetProperty("TargetState", out var state))
            {
                return state.GetString() ?? "Unknown";
            }
        }
        catch { }

        return "Unknown";
    }

    private async void RunButton_Click(object? sender, EventArgs e)
    {
        if (_cancellationTokenSource != null)
        {
            // Cancel running task
            _cancellationTokenSource.Cancel();
            _runButton.Text = "\u25B6 Run Task";
            _runButton.BackColor = Color.FromArgb(0, 120, 212);
            return;
        }

        var connection = await _connectionManager.GetConnectionAsync(_task.ConnectionId);
        if (connection == null)
        {
            MessageBox.Show("Connection not found.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Get cookies
        var cookies = _authService.GetStoredCookies(connection.AdminDomain);
        if (cookies == null || !cookies.IsValid)
        {
            MessageBox.Show("Please authenticate first.", "Authentication Required",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _cancellationTokenSource = new CancellationTokenSource();
        _runButton.Text = "\u23F9 Cancel";
        _runButton.BackColor = Color.FromArgb(220, 53, 69);
        _progressPanel.Visible = true;
        _deleteButton.Enabled = false;

        var targetState = GetTargetState();
        var lockState = targetState switch
        {
            "Unlock (Active)" => "Unlock",
            "Read Only" => "ReadOnly",
            "No Access (Restricted)" => "NoAccess",
            _ => "Unlock"
        };

        try
        {
            _task.Status = Models.TaskStatus.Running;
            _task.LastRunAt = DateTime.UtcNow;
            await _taskService.SaveTaskAsync(_task);

            using var spService = new SharePointService(cookies, connection.AdminDomain);

            var successCount = 0;
            var failCount = 0;

            for (int i = 0; i < _task.TargetSiteUrls.Count; i++)
            {
                if (_cancellationTokenSource.Token.IsCancellationRequested)
                    break;

                var siteUrl = _task.TargetSiteUrls[i];
                _progressBar.Value = (int)((i + 1) * 100.0 / _task.TargetSiteUrls.Count);
                _progressLabel.Text = $"Processing {i + 1} of {_task.TargetSiteUrls.Count}: {siteUrl}";

                try
                {
                    var result = await spService.SetSiteLockStateAsync(siteUrl, lockState);

                    if (result.IsSuccess)
                    {
                        _resultsGrid.Rows[i].Cells["Status"].Value = "Success";
                        _resultsGrid.Rows[i].Cells["NewState"].Value = targetState;
                        _resultsGrid.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(212, 237, 218);
                        successCount++;
                    }
                    else
                    {
                        _resultsGrid.Rows[i].Cells["Status"].Value = "Failed";
                        _resultsGrid.Rows[i].Cells["Error"].Value = result.ErrorMessage;
                        _resultsGrid.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(248, 215, 218);
                        failCount++;
                    }
                }
                catch (Exception ex)
                {
                    _resultsGrid.Rows[i].Cells["Status"].Value = "Error";
                    _resultsGrid.Rows[i].Cells["Error"].Value = ex.Message;
                    _resultsGrid.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(248, 215, 218);
                    failCount++;
                }
            }

            _task.Status = failCount == 0 ? Models.TaskStatus.Completed : Models.TaskStatus.Failed;
            await _taskService.SaveTaskAsync(_task);

            _progressLabel.Text = $"Completed: {successCount} succeeded, {failCount} failed";

            // Build message with error details if any failed
            var message = $"Task completed.\n\nSucceeded: {successCount}\nFailed: {failCount}";
            if (failCount > 0)
            {
                // Get first few error messages from the grid
                var errors = new List<string>();
                foreach (DataGridViewRow row in _resultsGrid.Rows)
                {
                    if (row.Cells["Status"].Value?.ToString() == "Failed" ||
                        row.Cells["Status"].Value?.ToString() == "Error")
                    {
                        var error = row.Cells["Error"].Value?.ToString();
                        if (!string.IsNullOrEmpty(error))
                        {
                            errors.Add($"- {row.Cells["SiteUrl"].Value}: {error}");
                            if (errors.Count >= 3) break;
                        }
                    }
                }
                if (errors.Count > 0)
                {
                    message += "\n\nErrors:\n" + string.Join("\n", errors);
                    if (failCount > errors.Count)
                    {
                        message += $"\n... and {failCount - errors.Count} more";
                    }
                }
            }

            MessageBox.Show(message, "Task Complete", MessageBoxButtons.OK,
                failCount == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }
        catch (OperationCanceledException)
        {
            _progressLabel.Text = "Cancelled";
            _task.Status = Models.TaskStatus.Pending;
            await _taskService.SaveTaskAsync(_task);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Task failed: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            _task.Status = Models.TaskStatus.Failed;
            await _taskService.SaveTaskAsync(_task);
        }
        finally
        {
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
            _runButton.Text = "\u25B6 Run Task";
            _runButton.BackColor = Color.FromArgb(0, 120, 212);
            _deleteButton.Enabled = true;
            UpdateUI();
        }
    }

    private async void DeleteButton_Click(object? sender, EventArgs e)
    {
        var result = MessageBox.Show(
            $"Are you sure you want to delete task '{_task.Name}'?",
            "Delete Task",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (result == DialogResult.Yes)
        {
            await _taskService.DeleteTaskAsync(_task.Id);
            await NavigationService!.GoBackAsync();
        }
    }
}

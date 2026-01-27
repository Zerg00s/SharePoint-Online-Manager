using System.Text.Json;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for executing and viewing Remove Site Collection Administrators task results.
/// </summary>
public class RemoveSiteAdminsDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Label _adminsLabel = null!;
    private Button _runButton = null!;
    private Button _deleteButton = null!;
    private DataGridView _resultsGrid = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private Panel _progressPanel = null!;

    private TaskDefinition _task = null!;
    private AddSiteAdminsConfiguration? _config;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;

    public override string ScreenTitle => _task?.Name ?? "Remove Site Admins";

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

        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 120
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

        _adminsLabel = new Label
        {
            Location = new Point(0, 55),
            AutoSize = true,
            ForeColor = Color.DarkRed
        };

        var buttonPanel = new FlowLayoutPanel
        {
            Location = new Point(0, 85),
            Size = new Size(700, 35),
            FlowDirection = FlowDirection.LeftToRight
        };

        _runButton = new Button
        {
            Text = "\u25B6 Run Task",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(220, 53, 69),
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
        headerPanel.Controls.AddRange([_taskNameLabel, _taskInfoLabel, _adminsLabel, buttonPanel]);

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
        _resultsGrid.Columns.Add("Admin", "Admin Removed");
        _resultsGrid.Columns.Add("Status", "Status");
        _resultsGrid.Columns.Add("Error", "Error");

        _resultsGrid.Columns["SiteUrl"].FillWeight = 30;
        _resultsGrid.Columns["SiteTitle"].FillWeight = 20;
        _resultsGrid.Columns["Admin"].FillWeight = 25;
        _resultsGrid.Columns["Status"].FillWeight = 10;
        _resultsGrid.Columns["Error"].FillWeight = 15;

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
        else if (parameter is TaskExecutionParameter execParam)
        {
            _task = execParam.Task;
            if (execParam.ExecuteImmediately)
            {
                BeginInvoke(() => RunButton_Click(null, EventArgs.Empty));
            }
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

        if (!string.IsNullOrEmpty(_task.ConfigurationJson))
        {
            try
            {
                _config = JsonSerializer.Deserialize<AddSiteAdminsConfiguration>(_task.ConfigurationJson);
            }
            catch
            {
                _config = null;
            }
        }

        UpdateUI();
    }

    private void UpdateUI()
    {
        _taskNameLabel.Text = _task.Name;

        var adminCount = _config?.Administrators.Count ?? 0;
        _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Sites: {_task.TargetSiteUrls.Count} | " +
                              $"Administrators: {adminCount} | Status: {_task.Status}";

        if (_config?.Administrators.Count > 0)
        {
            var adminNames = string.Join(", ", _config.Administrators.Select(a => a.DisplayName));
            _adminsLabel.Text = $"Removing: {adminNames}";
        }
        else
        {
            _adminsLabel.Text = "No administrators configured";
        }

        UpdateTitle();

        _resultsGrid.Rows.Clear();
        foreach (var siteUrl in _task.TargetSiteUrls)
        {
            _resultsGrid.Rows.Add(siteUrl, "", "", "Pending", "");
        }
    }

    private async void RunButton_Click(object? sender, EventArgs e)
    {
        if (_cancellationTokenSource != null)
        {
            _cancellationTokenSource.Cancel();
            _runButton.Text = "\u25B6 Run Task";
            _runButton.BackColor = Color.FromArgb(220, 53, 69);
            return;
        }

        if (_config?.Administrators.Count == 0)
        {
            MessageBox.Show("No administrators configured.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var connection = await _connectionManager.GetConnectionAsync(_task.ConnectionId);
        if (connection == null)
        {
            MessageBox.Show("Connection not found.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var cookies = _authService.GetStoredCookies(connection.AdminDomain);
        if (cookies == null || !cookies.IsValid)
        {
            MessageBox.Show("Please authenticate first.", "Authentication Required",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _cancellationTokenSource = new CancellationTokenSource();
        _runButton.Text = "\u23F9 Cancel";
        _runButton.BackColor = Color.Gray;
        _progressPanel.Visible = true;
        _deleteButton.Enabled = false;

        try
        {
            _task.Status = Models.TaskStatus.Running;
            _task.LastRunAt = DateTime.UtcNow;
            await _taskService.SaveTaskAsync(_task);

            using var spService = new SharePointService(cookies, connection.AdminDomain);

            var totalOperations = _task.TargetSiteUrls.Count * (_config?.Administrators.Count ?? 0);
            var currentOperation = 0;
            var successCount = 0;
            var failCount = 0;

            for (int siteIndex = 0; siteIndex < _task.TargetSiteUrls.Count; siteIndex++)
            {
                if (_cancellationTokenSource.Token.IsCancellationRequested)
                    break;

                var siteUrl = _task.TargetSiteUrls[siteIndex];
                var siteSuccess = true;
                var siteErrors = new List<string>();
                var removedAdmins = new List<string>();

                foreach (var admin in _config?.Administrators ?? [])
                {
                    if (_cancellationTokenSource.Token.IsCancellationRequested)
                        break;

                    currentOperation++;
                    _progressBar.Value = (int)(currentOperation * 100.0 / totalOperations);
                    _progressLabel.Text = $"Processing {currentOperation} of {totalOperations}: {siteUrl} - {admin.DisplayName}";

                    try
                    {
                        var result = await spService.RemoveSiteCollectionAdminAsync(siteUrl, admin.LoginName);

                        if (result.IsSuccess)
                        {
                            removedAdmins.Add(admin.DisplayName);
                        }
                        else
                        {
                            siteSuccess = false;
                            siteErrors.Add($"{admin.DisplayName}: {result.ErrorMessage}");
                        }
                    }
                    catch (Exception ex)
                    {
                        siteSuccess = false;
                        siteErrors.Add($"{admin.DisplayName}: {ex.Message}");
                    }
                }

                _resultsGrid.Rows[siteIndex].Cells["Admin"].Value = string.Join(", ", removedAdmins);

                if (siteSuccess && removedAdmins.Count > 0)
                {
                    _resultsGrid.Rows[siteIndex].Cells["Status"].Value = "Success";
                    _resultsGrid.Rows[siteIndex].DefaultCellStyle.BackColor = Color.FromArgb(212, 237, 218);
                    successCount++;
                }
                else if (removedAdmins.Count > 0)
                {
                    _resultsGrid.Rows[siteIndex].Cells["Status"].Value = "Partial";
                    _resultsGrid.Rows[siteIndex].Cells["Error"].Value = string.Join("; ", siteErrors);
                    _resultsGrid.Rows[siteIndex].DefaultCellStyle.BackColor = Color.FromArgb(255, 243, 205);
                    successCount++;
                }
                else
                {
                    _resultsGrid.Rows[siteIndex].Cells["Status"].Value = "Failed";
                    _resultsGrid.Rows[siteIndex].Cells["Error"].Value = string.Join("; ", siteErrors);
                    _resultsGrid.Rows[siteIndex].DefaultCellStyle.BackColor = Color.FromArgb(248, 215, 218);
                    failCount++;
                }
            }

            _task.Status = failCount == 0 ? Models.TaskStatus.Completed : Models.TaskStatus.Failed;
            await _taskService.SaveTaskAsync(_task);

            _progressLabel.Text = $"Completed: {successCount} succeeded, {failCount} failed";

            var message = $"Task completed.\n\nSucceeded: {successCount}\nFailed: {failCount}";
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
            _runButton.BackColor = Color.FromArgb(220, 53, 69);
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

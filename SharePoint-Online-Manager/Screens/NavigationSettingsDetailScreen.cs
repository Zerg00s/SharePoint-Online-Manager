using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing navigation settings comparison results and applying settings.
/// </summary>
public class NavigationSettingsDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runCompareButton = null!;
    private Button _applySettingsButton = null!;
    private Button _exportButton = null!;
    private Button _deleteButton = null!;
    private DataGridView _resultsGrid = null!;
    private TextBox _logTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private ComboBox _filterCombo = null!;

    private TaskDefinition _task = null!;
    private NavigationSettingsResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private CsvExporter _csvExporter = null!;

    public override string ScreenTitle => _task?.Name ?? "Navigation Settings";

    protected override void OnInitialize()
    {
        _taskService = GetRequiredService<ITaskService>();
        _authService = GetRequiredService<IAuthenticationService>();
        _connectionManager = GetRequiredService<IConnectionManager>();
        _csvExporter = GetRequiredService<CsvExporter>();
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
            Size = new Size(900, 35),
            FlowDirection = FlowDirection.LeftToRight
        };

        _runCompareButton = new Button
        {
            Text = "\U0001F504 Compare",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _runCompareButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _runCompareButton.FlatAppearance.BorderSize = 1;
        _runCompareButton.Click += RunCompareButton_Click;

        _applySettingsButton = new Button
        {
            Text = "\u2705 Apply to Target",
            Size = new Size(130, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 150, 0),
            ForeColor = Color.White
        };
        _applySettingsButton.FlatAppearance.BorderColor = Color.FromArgb(0, 150, 0);
        _applySettingsButton.FlatAppearance.BorderSize = 1;
        _applySettingsButton.Click += ApplySettingsButton_Click;

        _exportButton = new Button
        {
            Text = "\U0001F4BE Export",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportButton.FlatAppearance.BorderSize = 1;
        _exportButton.Click += ExportButton_Click;

        _deleteButton = new Button
        {
            Text = "\U0001F5D1 Delete Task",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            ForeColor = Color.DarkRed
        };
        _deleteButton.FlatAppearance.BorderColor = Color.DarkRed;
        _deleteButton.FlatAppearance.BorderSize = 1;
        _deleteButton.Click += DeleteButton_Click;

        buttonPanel.Controls.AddRange(new Control[] { _runCompareButton, _applySettingsButton, _exportButton, _deleteButton });

        headerPanel.Controls.AddRange(new Control[] { _taskNameLabel, _taskInfoLabel, buttonPanel });

        // Progress panel
        var progressPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 50,
            Visible = false,
            Name = "ProgressPanel"
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
        var resultsTab = new TabPage("Comparison Results");

        var resultsHeaderPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 35,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        var filterLabel = new Label
        {
            Text = "Filter:",
            AutoSize = true,
            Padding = new Padding(0, 5, 5, 0)
        };

        _filterCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 180
        };
        _filterCombo.Items.AddRange(new object[]
        {
            "All",
            "Mismatches Only",
            "Matches Only",
            "Errors Only"
        });
        _filterCombo.SelectedIndex = 0;
        _filterCombo.SelectedIndexChanged += FilterCombo_SelectedIndexChanged;

        resultsHeaderPanel.Controls.AddRange(new Control[] { filterLabel, _filterCombo });

        _resultsGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            MultiSelect = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        };

        // Add context menu for copying
        var gridContextMenu = new ContextMenuStrip();
        var copyMenuItem = new ToolStripMenuItem("Copy", null, (s, e) =>
        {
            if (_resultsGrid.CurrentCell != null)
            {
                Clipboard.SetText(_resultsGrid.CurrentCell.Value?.ToString() ?? "");
            }
        });
        copyMenuItem.ShortcutKeys = Keys.Control | Keys.C;
        gridContextMenu.Items.Add(copyMenuItem);
        _resultsGrid.ContextMenuStrip = gridContextMenu;

        _resultsGrid.Columns.Add("SourceSite", "Source Site");
        _resultsGrid.Columns.Add("TargetSite", "Target Site");
        _resultsGrid.Columns.Add("SourceHNav", "Source Horizontal Nav");
        _resultsGrid.Columns.Add("TargetHNav", "Target Horizontal Nav");
        _resultsGrid.Columns.Add("HNavMatch", "Horizontal Nav Match");
        _resultsGrid.Columns.Add("SourceMM", "Source Mega Menu");
        _resultsGrid.Columns.Add("TargetMM", "Target Mega Menu");
        _resultsGrid.Columns.Add("MMMatch", "Mega Menu Match");
        _resultsGrid.Columns.Add("Status", "Status");

        // Adjust column widths
        _resultsGrid.Columns["SourceSite"].FillWeight = 150;
        _resultsGrid.Columns["TargetSite"].FillWeight = 150;
        _resultsGrid.Columns["SourceHNav"].FillWeight = 70;
        _resultsGrid.Columns["TargetHNav"].FillWeight = 70;
        _resultsGrid.Columns["HNavMatch"].FillWeight = 70;
        _resultsGrid.Columns["SourceMM"].FillWeight = 70;
        _resultsGrid.Columns["TargetMM"].FillWeight = 70;
        _resultsGrid.Columns["MMMatch"].FillWeight = 70;
        _resultsGrid.Columns["Status"].FillWeight = 60;

        resultsTab.Controls.Add(_resultsGrid);
        resultsTab.Controls.Add(resultsHeaderPanel);

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

        _tabControl.TabPages.AddRange(new[] { resultsTab, logTab });

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

        if (task == null && _task != null)
        {
            await RefreshTaskDetailsAsync();
            return;
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
            await ExecuteCompareAsync();
        }
    }

    private async Task RefreshTaskDetailsAsync()
    {
        _taskNameLabel.Text = _task.Name;

        var connection = await _connectionManager.GetConnectionAsync(_task.ConnectionId);
        var connectionName = connection?.Name ?? "Unknown";

        _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Status: {_task.StatusDescription} | " +
                              $"Connection: {connectionName}";

        // Load latest result
        _currentResult = await _taskService.GetLatestNavigationSettingsResultAsync(_task.Id);

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportButton.Enabled = true;
            _applySettingsButton.Enabled = _currentResult.GetMismatchedSites().Any() && !_currentResult.ApplyMode;
        }
        else
        {
            _resultsGrid.Rows.Clear();
            _logTextBox.Text = "No results yet. Click 'Compare' to compare navigation settings.";
            _exportButton.Enabled = false;
            _applySettingsButton.Enabled = false;
        }
    }

    private void DisplayResults(NavigationSettingsResult result)
    {
        DisplayComparisonResults(result);
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);
    }

    private void DisplayComparisonResults(NavigationSettingsResult result)
    {
        _resultsGrid.Rows.Clear();

        var items = result.SiteResults.ToList();

        // Apply filter
        var filter = _filterCombo.SelectedIndex;
        items = filter switch
        {
            1 => items.Where(i => i.Status == NavigationSettingsStatus.Mismatch).ToList(),
            2 => items.Where(i => i.Status == NavigationSettingsStatus.Match || i.Status == NavigationSettingsStatus.Applied).ToList(),
            3 => items.Where(i => i.Status == NavigationSettingsStatus.Error || i.Status == NavigationSettingsStatus.Failed).ToList(),
            _ => items
        };

        foreach (var item in items)
        {
            var rowIndex = _resultsGrid.Rows.Add(
                item.SourceSiteUrl,
                item.TargetSiteUrl,
                item.SourceHorizontalQuickLaunch ? "Yes" : "No",
                item.TargetHorizontalQuickLaunch ? "Yes" : "No",
                item.HorizontalQuickLaunchMatches ? "Yes" : "No",
                item.SourceMegaMenuEnabled ? "Yes" : "No",
                item.TargetMegaMenuEnabled ? "Yes" : "No",
                item.MegaMenuEnabledMatches ? "Yes" : "No",
                item.StatusDescription
            );

            // Color coding
            var row = _resultsGrid.Rows[rowIndex];
            row.DefaultCellStyle.BackColor = item.Status switch
            {
                NavigationSettingsStatus.Match => Color.FromArgb(200, 255, 200), // Green
                NavigationSettingsStatus.Applied => Color.FromArgb(180, 220, 255), // Light blue
                NavigationSettingsStatus.Mismatch => Color.FromArgb(255, 255, 150), // Yellow
                NavigationSettingsStatus.Error => Color.FromArgb(255, 200, 200), // Red
                NavigationSettingsStatus.Failed => Color.FromArgb(255, 200, 200), // Red
                _ => SystemColors.Window
            };

            // Color individual match cells
            if (!item.HorizontalQuickLaunchMatches)
            {
                row.Cells["HNavMatch"].Style.BackColor = Color.FromArgb(255, 200, 200);
            }
            if (!item.MegaMenuEnabledMatches)
            {
                row.Cells["MMMatch"].Style.BackColor = Color.FromArgb(255, 200, 200);
            }
        }

        SetStatus($"Showing {items.Count} site comparison(s)");
    }

    private void FilterCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayComparisonResults(_currentResult);
        }
    }

    private async void RunCompareButton_Click(object? sender, EventArgs e)
    {
        if (_runCompareButton.Text == "Cancel")
        {
            _cancellationTokenSource?.Cancel();
            return;
        }

        await ExecuteCompareAsync();
    }

    private async Task ExecuteCompareAsync()
    {
        await ExecuteTaskAsync(applyMode: false);
    }

    private async void ApplySettingsButton_Click(object? sender, EventArgs e)
    {
        var mismatchCount = _currentResult?.GetMismatchedSites().Count() ?? 0;

        var result = MessageBox.Show(
            $"This will apply navigation settings from source to {mismatchCount} target site(s).\n\n" +
            "Settings to be applied:\n" +
            "- HorizontalQuickLaunch\n" +
            "- MegaMenuEnabled\n\n" +
            "Are you sure you want to continue?",
            "Apply Navigation Settings",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (result != DialogResult.Yes)
            return;

        await ExecuteTaskAsync(applyMode: true);
    }

    private async Task ExecuteTaskAsync(bool applyMode)
    {
        // Setup for execution
        _runCompareButton.Text = "Cancel";
        _deleteButton.Enabled = false;
        _exportButton.Enabled = false;
        _applySettingsButton.Enabled = false;

        var progressPanel = Controls.Find("ProgressPanel", false).FirstOrDefault();
        if (progressPanel != null)
        {
            progressPanel.Visible = true;
        }
        _progressBar.Value = 0;
        _progressLabel.Text = applyMode ? "Applying settings..." : "Comparing...";

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
            _currentResult = await _taskService.ExecuteNavigationSettingsSyncAsync(
                _task,
                _authService,
                _connectionManager,
                applyMode,
                progress,
                _cancellationTokenSource.Token);

            DisplayResults(_currentResult);

            if (_currentResult.Success)
            {
                if (applyMode)
                {
                    SetStatus($"Settings applied successfully. Applied: {_currentResult.AppliedPairs}, Failed: {_currentResult.FailedPairs}");
                }
                else
                {
                    SetStatus($"Comparison complete. Matches: {_currentResult.MatchingPairs}, Mismatches: {_currentResult.MismatchedPairs}");
                }
            }
            else
            {
                SetStatus($"Task completed with errors: {_currentResult.ErrorMessage}");
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
            _runCompareButton.Text = "\U0001F504 Compare";
            _deleteButton.Enabled = true;
            _exportButton.Enabled = _currentResult != null;
            _applySettingsButton.Enabled = _currentResult?.GetMismatchedSites().Any() == true && !(_currentResult?.ApplyMode ?? false);
            if (progressPanel != null)
            {
                progressPanel.Visible = false;
            }
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;

            // Refresh task details
            var updatedTask = await _taskService.GetTaskAsync(_task.Id);
            if (updatedTask != null)
            {
                _task = updatedTask;
                _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Status: {_task.StatusDescription}";
            }
        }
    }

    private void ExportButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_NavigationSettings_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportNavigationSettingsReport(_currentResult, dialog.FileName);
                SetStatus($"Exported to {dialog.FileName}");
                OfferToOpenFile(dialog.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private static void OfferToOpenFile(string filePath)
    {
        var result = MessageBox.Show(
            "Export completed. Would you like to open the file?",
            "Export Complete",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Information);

        if (result == DialogResult.Yes)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            });
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

    private static string SanitizeFileName(string fileName)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        foreach (var c in invalidChars)
        {
            fileName = fileName.Replace(c, '_');
        }
        return fileName;
    }
}

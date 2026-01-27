using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing list compare task details and results.
/// </summary>
public class ListCompareDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _exportAllButton = null!;
    private Button _exportIssuesButton = null!;
    private Button _exportMappingButton = null!;
    private Button _deleteButton = null!;
    private DataGridView _resultsGrid = null!;
    private ListView _issuesList = null!;
    private TextBox _logTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private ComboBox _filterCombo = null!;

    private TaskDefinition _task = null!;
    private ListCompareResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private CsvExporter _csvExporter = null!;

    public override string ScreenTitle => _task?.Name ?? "List Compare Details";

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

        _runButton = new Button
        {
            Text = "\u25B6 Run Task",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _runButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _runButton.FlatAppearance.BorderSize = 1;
        _runButton.Click += RunButton_Click;

        _exportAllButton = new Button
        {
            Text = "\U0001F4BE Export All",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportAllButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportAllButton.FlatAppearance.BorderSize = 1;
        _exportAllButton.Click += ExportAllButton_Click;

        _exportIssuesButton = new Button
        {
            Text = "\u26A0 Export Issues",
            Size = new Size(120, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportIssuesButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportIssuesButton.FlatAppearance.BorderSize = 1;
        _exportIssuesButton.Click += ExportIssuesButton_Click;

        _exportMappingButton = new Button
        {
            Text = "\U0001F4CB Export Mapping",
            Size = new Size(140, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportMappingButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportMappingButton.FlatAppearance.BorderSize = 1;
        _exportMappingButton.Click += ExportMappingButton_Click;

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

        buttonPanel.Controls.AddRange(new Control[] { _runButton, _exportAllButton, _exportIssuesButton, _exportMappingButton, _deleteButton });

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

        // Tab control for results, issues, and log
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
            Width = 220
        };
        _filterCombo.Items.AddRange(new object[]
        {
            "All",
            "Mismatches Only",
            "Exceeds Threshold",
            "Missing on Target",
            "Missing on Source",
            "Under-migrated (Source > Target)"
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

        // Add right-click context menu for copying
        var gridContextMenu = new ContextMenuStrip();
        var copyMenuItem = new ToolStripMenuItem("Copy", null, (s, e) =>
        {
            if (_resultsGrid.CurrentCell != null)
            {
                Clipboard.SetText(_resultsGrid.CurrentCell.Value?.ToString() ?? "");
            }
        });
        copyMenuItem.ShortcutKeys = Keys.Control | Keys.C;
        var copyRowMenuItem = new ToolStripMenuItem("Copy Row", null, (s, e) =>
        {
            if (_resultsGrid.CurrentRow != null)
            {
                var values = new List<string>();
                foreach (DataGridViewCell cell in _resultsGrid.CurrentRow.Cells)
                {
                    values.Add(cell.Value?.ToString() ?? "");
                }
                Clipboard.SetText(string.Join("\t", values));
            }
        });
        gridContextMenu.Items.Add(copyMenuItem);
        gridContextMenu.Items.Add(copyRowMenuItem);
        _resultsGrid.ContextMenuStrip = gridContextMenu;

        _resultsGrid.Columns.Add("SourceSite", "Source Site");
        _resultsGrid.Columns.Add("TargetSite", "Target Site");
        _resultsGrid.Columns.Add("ListTitle", "List Title");
        _resultsGrid.Columns.Add("ListType", "Type");
        _resultsGrid.Columns.Add("SourceCount", "Source Count");
        _resultsGrid.Columns.Add("TargetCount", "Target Count");
        _resultsGrid.Columns.Add("Difference", "Difference");
        _resultsGrid.Columns.Add("PercentDiff", "% Diff");
        _resultsGrid.Columns.Add("Status", "Status");

        // Adjust column widths
        _resultsGrid.Columns["SourceSite"].FillWeight = 150;
        _resultsGrid.Columns["TargetSite"].FillWeight = 150;
        _resultsGrid.Columns["ListTitle"].FillWeight = 100;
        _resultsGrid.Columns["ListType"].FillWeight = 60;
        _resultsGrid.Columns["SourceCount"].FillWeight = 50;
        _resultsGrid.Columns["TargetCount"].FillWeight = 50;
        _resultsGrid.Columns["Difference"].FillWeight = 50;
        _resultsGrid.Columns["PercentDiff"].FillWeight = 40;
        _resultsGrid.Columns["Status"].FillWeight = 60;

        resultsTab.Controls.Add(_resultsGrid);
        resultsTab.Controls.Add(resultsHeaderPanel);

        // Issues Summary tab
        var issuesTab = new TabPage("Issues Summary");
        _issuesList = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _issuesList.Columns.Add("Source Site", 250);
        _issuesList.Columns.Add("Target Site", 250);
        _issuesList.Columns.Add("Mismatches", 80);
        _issuesList.Columns.Add("Source Only", 80);
        _issuesList.Columns.Add("Target Only", 80);
        _issuesList.Columns.Add("Error", 200);

        issuesTab.Controls.Add(_issuesList);

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

        _tabControl.TabPages.AddRange(new[] { resultsTab, issuesTab, logTab });

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

        // If no parameter but we already have a task (back navigation), use existing task
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
            await ExecuteTaskAsync();
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
        _currentResult = await _taskService.GetLatestListCompareResultAsync(_task.Id);

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportAllButton.Enabled = true;
            _exportIssuesButton.Enabled = _currentResult.GetSitesWithIssues().Any();
            _exportMappingButton.Enabled = true;
        }
        else
        {
            _resultsGrid.Rows.Clear();
            _issuesList.Items.Clear();
            _logTextBox.Text = "No results yet. Click 'Run Task' to execute.";
            _exportAllButton.Enabled = false;
            _exportIssuesButton.Enabled = false;
            _exportMappingButton.Enabled = false;
        }
    }

    private void DisplayResults(ListCompareResult result)
    {
        DisplayComparisonResults(result);
        DisplayIssuesSummary(result);
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);
    }

    private void DisplayComparisonResults(ListCompareResult result)
    {
        _resultsGrid.Rows.Clear();

        var items = result.GetAllListComparisons().ToList();

        // Apply filter
        var filter = _filterCombo.SelectedIndex;
        items = filter switch
        {
            1 => items.Where(i => i.Status == ListCompareStatus.Mismatch).ToList(),
            2 => items.Where(i => i.Status == ListCompareStatus.Mismatch ||
                                  i.Status == ListCompareStatus.SourceOnly ||
                                  i.Status == ListCompareStatus.TargetOnly).ToList(),
            3 => items.Where(i => i.Status == ListCompareStatus.SourceOnly).ToList(),
            4 => items.Where(i => i.Status == ListCompareStatus.TargetOnly).ToList(),
            5 => items.Where(i => i.Status == ListCompareStatus.SourceOnly ||
                                  (i.Status == ListCompareStatus.Mismatch && i.SourceCount > i.TargetCount)).ToList(),
            _ => items
        };

        foreach (var item in items)
        {
            var rowIndex = _resultsGrid.Rows.Add(
                item.SourceSiteUrl,
                item.TargetSiteUrl,
                item.ListTitle,
                item.ListType,
                item.SourceCount,
                item.TargetCount,
                item.Difference.ToString("+0;-0;0"),
                $"{item.PercentDifference:F1}%",
                item.StatusDescription
            );

            // Color coding
            var row = _resultsGrid.Rows[rowIndex];
            row.DefaultCellStyle.BackColor = item.Status switch
            {
                ListCompareStatus.Match => Color.FromArgb(200, 255, 200), // Green
                ListCompareStatus.Mismatch => Color.FromArgb(255, 255, 150), // Yellow
                ListCompareStatus.SourceOnly => Color.FromArgb(255, 200, 200), // Red
                ListCompareStatus.TargetOnly => Color.FromArgb(255, 220, 180), // Orange
                _ => SystemColors.Window
            };
        }

        SetStatus($"Showing {items.Count} list comparison(s)");
    }

    private void DisplayIssuesSummary(ListCompareResult result)
    {
        _issuesList.Items.Clear();

        foreach (var site in result.GetSitesWithIssues())
        {
            var item = new ListViewItem(site.SourceSiteUrl);
            item.SubItems.Add(site.TargetSiteUrl);
            item.SubItems.Add(site.MismatchCount.ToString());
            item.SubItems.Add(site.SourceOnlyCount.ToString());
            item.SubItems.Add(site.TargetOnlyCount.ToString());
            item.SubItems.Add(site.ErrorMessage ?? "");

            // Color by severity
            if (!site.Success)
            {
                item.BackColor = Color.FromArgb(255, 200, 200);
            }
            else if (site.MismatchCount > 0 || site.SourceOnlyCount > 0 || site.TargetOnlyCount > 0)
            {
                item.BackColor = Color.FromArgb(255, 255, 150);
            }

            _issuesList.Items.Add(item);
        }
    }

    private void FilterCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayComparisonResults(_currentResult);
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
        // Setup for execution
        _runButton.Text = "Cancel";
        _deleteButton.Enabled = false;
        _exportAllButton.Enabled = false;
        _exportIssuesButton.Enabled = false;
        _exportMappingButton.Enabled = false;

        var progressPanel = Controls.Find("ProgressPanel", false).FirstOrDefault();
        if (progressPanel != null)
        {
            progressPanel.Visible = true;
        }
        _progressBar.Value = 0;
        _progressLabel.Text = "Starting...";

        _cancellationTokenSource = new CancellationTokenSource();
        _resultsGrid.Rows.Clear();
        _issuesList.Items.Clear();
        _logTextBox.Clear();

        var progress = new Progress<TaskProgress>(p =>
        {
            _progressBar.Value = p.PercentComplete;
            _progressLabel.Text = p.Message;
        });

        try
        {
            _currentResult = await _taskService.ExecuteListCompareAsync(
                _task,
                _authService,
                _connectionManager,
                progress,
                _cancellationTokenSource.Token);

            DisplayResults(_currentResult);

            if (_currentResult.Success)
            {
                SetStatus($"Task completed successfully. Processed {_currentResult.TotalPairsProcessed} site pairs.");
            }
            else
            {
                SetStatus($"Task completed with errors. {_currentResult.FailedPairs} site pair(s) failed.");
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
            _exportAllButton.Enabled = _currentResult != null;
            _exportIssuesButton.Enabled = _currentResult?.GetSitesWithIssues().Any() ?? false;
            _exportMappingButton.Enabled = _currentResult != null;
            if (progressPanel != null)
            {
                progressPanel.Visible = false;
            }
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;

            // Refresh task details to show updated status
            var updatedTask = await _taskService.GetTaskAsync(_task.Id);
            if (updatedTask != null)
            {
                _task = updatedTask;
                _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Status: {_task.StatusDescription}";
            }
        }
    }

    private void ExportAllButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_AllResults_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportListCompareReport(_currentResult, dialog.FileName);
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

    private void ExportIssuesButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_Issues_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportIssuesSummary(_currentResult, dialog.FileName);
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

    private void ExportMappingButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_ListMapping_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportListMapping(_currentResult, dialog.FileName);
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

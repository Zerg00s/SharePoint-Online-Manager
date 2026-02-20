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
    private TextBox _issuesFilterTextBox = null!;
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

        // Filter panel at top of issues tab
        var issuesFilterPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 35
        };
        var issuesFilterLabel = new Label
        {
            Text = "Filter:",
            AutoSize = true,
            Location = new Point(5, 10)
        };
        _issuesFilterTextBox = new TextBox
        {
            Location = new Point(50, 7),
            Size = new Size(300, 23),
            PlaceholderText = "Type to filter by URL..."
        };
        _issuesFilterTextBox.TextChanged += (s, e) =>
        {
            if (_currentResult != null)
                DisplayIssuesSummary(_currentResult);
        };
        issuesFilterPanel.Controls.Add(issuesFilterLabel);
        issuesFilterPanel.Controls.Add(_issuesFilterTextBox);

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
        _issuesList.ContextMenuStrip = CreateIssuesContextMenu();
        EnableCellTextSelection(_issuesList);
        EnableColumnSorting(_issuesList);

        issuesTab.Controls.Add(_issuesList);
        issuesTab.Controls.Add(issuesFilterPanel);

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

        var filterText = _issuesFilterTextBox?.Text?.Trim() ?? "";
        var sites = result.GetSitesWithIssues();
        if (!string.IsNullOrEmpty(filterText))
        {
            sites = sites.Where(s =>
                s.SourceSiteUrl.Contains(filterText, StringComparison.OrdinalIgnoreCase) ||
                s.TargetSiteUrl.Contains(filterText, StringComparison.OrdinalIgnoreCase));
        }

        foreach (var site in sites)
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

    /// <summary>
    /// Creates a context menu for the Issues list with copy items and export option.
    /// </summary>
    private ContextMenuStrip CreateIssuesContextMenu()
    {
        var contextMenu = new ContextMenuStrip();

        var copyCell = new ToolStripMenuItem("Copy Cell");
        copyCell.Click += (s, e) =>
        {
            if (_issuesList.SelectedItems.Count > 0)
            {
                var point = _issuesList.PointToClient(Cursor.Position);
                var hitTest = _issuesList.HitTest(point);
                if (hitTest.SubItem != null)
                {
                    var text = hitTest.SubItem.Text;
                    if (!string.IsNullOrEmpty(text))
                    {
                        Clipboard.SetText(text);
                    }
                }
            }
        };

        var copyRow = new ToolStripMenuItem("Copy Row");
        copyRow.Click += (s, e) =>
        {
            if (_issuesList.SelectedItems.Count > 0)
            {
                var item = _issuesList.SelectedItems[0];
                var values = new List<string>();
                for (int i = 0; i < item.SubItems.Count; i++)
                {
                    values.Add(item.SubItems[i].Text ?? "");
                }
                Clipboard.SetText(string.Join("\t", values));
            }
        };

        var copyAllUrls = new ToolStripMenuItem("Copy All URLs");
        copyAllUrls.Click += (s, e) =>
        {
            var urls = new List<string>();
            foreach (ListViewItem item in _issuesList.Items)
            {
                for (int i = 0; i < Math.Min(2, item.SubItems.Count); i++)
                {
                    var text = item.SubItems[i].Text ?? "";
                    if (text.StartsWith("http", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains(".sharepoint.com", StringComparison.OrdinalIgnoreCase))
                    {
                        urls.Add(text);
                    }
                }
            }
            if (urls.Count > 0)
            {
                Clipboard.SetText(string.Join(Environment.NewLine, urls));
            }
        };

        var exportIssuesCsv = new ToolStripMenuItem("Export Issues to CSV");
        exportIssuesCsv.Click += ExportIssuesToCsv_Click;

        contextMenu.Items.Add(copyCell);
        contextMenu.Items.Add(copyRow);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(copyAllUrls);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(exportIssuesCsv);

        contextMenu.Opening += (s, e) =>
        {
            exportIssuesCsv.Enabled = _currentResult != null;
        };

        return contextMenu;
    }

    private void ExportIssuesToCsv_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var answer = MessageBox.Show(
            "Include matching lists from these sites?\n\n" +
            "Yes = all lists from sites with issues\n" +
            "No = only issue rows (Mismatch, Source Only, Target Only)",
            "Export Issues",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (answer == DialogResult.Cancel)
            return;

        var issuesSites = _currentResult.GetSitesWithIssues().ToList();
        IEnumerable<ListCompareItem> items;

        if (answer == DialogResult.Yes)
        {
            items = issuesSites.SelectMany(s => s.ListComparisons);
        }
        else
        {
            items = issuesSites.SelectMany(s => s.ListComparisons)
                .Where(c => c.Status == ListCompareStatus.Mismatch ||
                            c.Status == ListCompareStatus.SourceOnly ||
                            c.Status == ListCompareStatus.TargetOnly);
        }

        var safeName = SanitizeFileName(_task.Name);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var defaultFileName = $"{safeName}_Issues_{timestamp}.csv";

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = defaultFileName
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportListCompareItems(items, dialog.FileName);
                SetStatus($"Exported issues to {dialog.FileName}");
                OfferToOpenFile(dialog.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    /// <summary>
    /// Enables click-to-select text on ListView cells by showing a read-only TextBox overlay.
    /// </summary>
    private static void EnableCellTextSelection(ListView listView)
    {
        TextBox? activeEditor = null;

        void DismissEditor()
        {
            var editor = activeEditor;
            activeEditor = null;
            if (editor != null && !editor.IsDisposed && !listView.IsDisposed)
            {
                listView.Controls.Remove(editor);
                editor.Dispose();
            }
        }

        listView.MouseClick += (s, e) =>
        {
            var hitTest = listView.HitTest(e.Location);
            if (hitTest.SubItem == null || hitTest.Item == null)
            {
                DismissEditor();
                return;
            }

            var text = hitTest.SubItem.Text;
            if (string.IsNullOrEmpty(text))
            {
                DismissEditor();
                return;
            }

            DismissEditor();

            var bounds = hitTest.SubItem.Bounds;

            activeEditor = new TextBox
            {
                Text = text,
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                Location = bounds.Location,
                Size = new Size(Math.Max(bounds.Width, 100), bounds.Height),
                Font = listView.Font,
                BackColor = SystemColors.Info
            };

            activeEditor.SelectAll();

            activeEditor.LostFocus += (_, _) => DismissEditor();
            activeEditor.KeyDown += (_, ke) =>
            {
                if (ke.KeyCode == Keys.Escape || ke.KeyCode == Keys.Enter)
                {
                    DismissEditor();
                    ke.Handled = true;
                }
            };

            listView.Controls.Add(activeEditor);
            activeEditor.Focus();
        };

        listView.ColumnWidthChanging += (s, e) => DismissEditor();
    }

    /// <summary>
    /// Enables column-click sorting on a ListView.
    /// </summary>
    private static void EnableColumnSorting(ListView listView)
    {
        int sortColumn = -1;
        SortOrder sortOrder = SortOrder.None;

        listView.ColumnClick += (s, e) =>
        {
            if (e.Column == sortColumn)
            {
                sortOrder = sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                sortColumn = e.Column;
                sortOrder = SortOrder.Ascending;
            }

            listView.ListViewItemSorter = new ListViewColumnComparer(sortColumn, sortOrder);
            listView.Sort();
        };
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

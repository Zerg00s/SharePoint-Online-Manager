using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing broken OneNote notebooks report and applying fixes.
/// </summary>
public class BrokenOneNoteDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _fixButton = null!;
    private Button _exportButton = null!;
    private Button _exportSummaryButton = null!;
    private Button _deleteButton = null!;
    private DataGridView _notebooksGrid = null!;
    private ListView _summaryList = null!;
    private TextBox _logTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private ComboBox _filterSiteCombo = null!;
    private ComboBox _filterStatusCombo = null!;
    private TextBox _searchTextBox = null!;

    private TaskDefinition _task = null!;
    private BrokenOneNoteReportResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private CsvExporter _csvExporter = null!;

    public override string ScreenTitle => _task?.Name ?? "Broken OneNote Report";

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

        _fixButton = new Button
        {
            Text = "\U0001F527 Fix Broken",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(200, 80, 0),
            ForeColor = Color.White
        };
        _fixButton.FlatAppearance.BorderColor = Color.FromArgb(200, 80, 0);
        _fixButton.FlatAppearance.BorderSize = 1;
        _fixButton.Click += FixButton_Click;

        _exportButton = new Button
        {
            Text = "\U0001F4BE Export CSV",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportButton.FlatAppearance.BorderSize = 1;
        _exportButton.Click += ExportButton_Click;

        _exportSummaryButton = new Button
        {
            Text = "\U0001F4CB Export Summary",
            Size = new Size(140, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportSummaryButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportSummaryButton.FlatAppearance.BorderSize = 1;
        _exportSummaryButton.Click += ExportSummaryButton_Click;

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

        buttonPanel.Controls.AddRange(new Control[] { _runButton, _fixButton, _exportButton, _exportSummaryButton, _deleteButton });

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

        // Tab control
        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        // Notebooks tab
        var notebooksTab = new TabPage("Notebooks");

        var notebooksHeaderPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 35,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        var filterSiteLabel = new Label
        {
            Text = "Site:",
            AutoSize = true,
            Padding = new Padding(0, 5, 5, 0)
        };

        _filterSiteCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 200
        };
        _filterSiteCombo.Items.Add("All Sites");
        _filterSiteCombo.SelectedIndex = 0;
        _filterSiteCombo.SelectedIndexChanged += FilterCombo_SelectedIndexChanged;

        var filterStatusLabel = new Label
        {
            Text = "Status:",
            AutoSize = true,
            Padding = new Padding(10, 5, 5, 0)
        };

        _filterStatusCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 120
        };
        _filterStatusCombo.Items.Add("All");
        _filterStatusCombo.Items.Add("Broken Only");
        _filterStatusCombo.Items.Add("Fixed");
        _filterStatusCombo.Items.Add("Healthy");
        _filterStatusCombo.SelectedIndex = 0;
        _filterStatusCombo.SelectedIndexChanged += FilterCombo_SelectedIndexChanged;

        var searchLabel = new Label
        {
            Text = "Search:",
            AutoSize = true,
            Padding = new Padding(10, 5, 5, 0)
        };

        _searchTextBox = new TextBox
        {
            Width = 150,
            PlaceholderText = "Notebook name..."
        };
        _searchTextBox.TextChanged += SearchTextBox_TextChanged;

        notebooksHeaderPanel.Controls.AddRange(new Control[]
        {
            filterSiteLabel, _filterSiteCombo,
            filterStatusLabel, _filterStatusCombo,
            searchLabel, _searchTextBox
        });

        _notebooksGrid = new DataGridView
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

        // Context menu
        var gridContextMenu = new ContextMenuStrip();
        var copyMenuItem = new ToolStripMenuItem("Copy Cell", null, (s, e) =>
        {
            if (_notebooksGrid.CurrentCell != null)
            {
                Clipboard.SetText(_notebooksGrid.CurrentCell.Value?.ToString() ?? "");
            }
        });
        copyMenuItem.ShortcutKeys = Keys.Control | Keys.C;
        var copyRowMenuItem = new ToolStripMenuItem("Copy Row", null, (s, e) =>
        {
            if (_notebooksGrid.CurrentRow != null)
            {
                var values = new List<string>();
                foreach (DataGridViewCell cell in _notebooksGrid.CurrentRow.Cells)
                {
                    values.Add(cell.Value?.ToString() ?? "");
                }
                Clipboard.SetText(string.Join("\t", values));
            }
        });
        gridContextMenu.Items.Add(copyMenuItem);
        gridContextMenu.Items.Add(copyRowMenuItem);
        _notebooksGrid.ContextMenuStrip = gridContextMenu;

        _notebooksGrid.Columns.Add("SiteUrl", "Site URL");
        _notebooksGrid.Columns.Add("Library", "Library");
        _notebooksGrid.Columns.Add("NotebookName", "Notebook Name");
        _notebooksGrid.Columns.Add("ItemId", "Item ID");
        _notebooksGrid.Columns.Add("Status", "Status");
        _notebooksGrid.Columns.Add("HtmlFileType", "HTML File Type");

        _notebooksGrid.Columns["SiteUrl"]!.FillWeight = 160;
        _notebooksGrid.Columns["Library"]!.FillWeight = 100;
        _notebooksGrid.Columns["NotebookName"]!.FillWeight = 120;
        _notebooksGrid.Columns["ItemId"]!.FillWeight = 50;
        _notebooksGrid.Columns["Status"]!.FillWeight = 60;
        _notebooksGrid.Columns["HtmlFileType"]!.FillWeight = 100;

        notebooksTab.Controls.Add(_notebooksGrid);
        notebooksTab.Controls.Add(notebooksHeaderPanel);

        // Site Summary tab
        var summaryTab = new TabPage("Site Summary");
        _summaryList = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _summaryList.Columns.Add("Site URL", 250);
        _summaryList.Columns.Add("Site Title", 130);
        _summaryList.Columns.Add("Total Notebooks", 100);
        _summaryList.Columns.Add("Broken", 70);
        _summaryList.Columns.Add("Fixed", 60);
        _summaryList.Columns.Add("Status", 70);
        _summaryList.Columns.Add("Error", 180);

        summaryTab.Controls.Add(_summaryList);

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

        _tabControl.TabPages.AddRange(new[] { notebooksTab, summaryTab, logTab });

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
            await ExecuteTaskAsync();
        }
    }

    private async Task RefreshTaskDetailsAsync()
    {
        _taskNameLabel.Text = _task.Name;

        var connection = await _connectionManager.GetConnectionAsync(_task.ConnectionId);
        var connectionName = connection?.Name ?? "Unknown";

        _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Status: {_task.StatusDescription} | " +
                              $"Connection: {connectionName} | Sites: {_task.TotalSites}";

        _currentResult = await _taskService.GetLatestBrokenOneNoteReportResultAsync(_task.Id);

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportButton.Enabled = true;
            _exportSummaryButton.Enabled = true;
            _fixButton.Enabled = _currentResult.TotalBroken > 0;
        }
        else
        {
            _notebooksGrid.Rows.Clear();
            _summaryList.Items.Clear();
            _logTextBox.Text = "No results yet. Click 'Run Task' to execute.";
            _exportButton.Enabled = false;
            _exportSummaryButton.Enabled = false;
            _fixButton.Enabled = false;
        }
    }

    private void DisplayResults(BrokenOneNoteReportResult result)
    {
        PopulateSiteFilter(result);
        DisplayNotebooks(result);
        DisplaySiteSummary(result);
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);
    }

    private void PopulateSiteFilter(BrokenOneNoteReportResult result)
    {
        _filterSiteCombo.Items.Clear();
        _filterSiteCombo.Items.Add("All Sites");

        foreach (var site in result.SiteResults.Where(s => s.Success))
        {
            _filterSiteCombo.Items.Add(site.SiteUrl);
        }

        _filterSiteCombo.SelectedIndex = 0;
    }

    private void DisplayNotebooks(BrokenOneNoteReportResult result)
    {
        _notebooksGrid.Rows.Clear();

        var notebooks = result.GetAllNotebooks().ToList();

        // Apply site filter
        var selectedSite = _filterSiteCombo.SelectedIndex > 0 ? _filterSiteCombo.SelectedItem?.ToString() : null;
        if (!string.IsNullOrEmpty(selectedSite))
        {
            notebooks = notebooks.Where(n => n.SiteUrl.Equals(selectedSite, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        // Apply status filter
        var statusFilter = _filterStatusCombo.SelectedItem?.ToString();
        if (statusFilter == "Broken Only")
        {
            notebooks = notebooks.Where(n => n.IsBroken && !n.IsFixed).ToList();
        }
        else if (statusFilter == "Fixed")
        {
            notebooks = notebooks.Where(n => n.IsFixed).ToList();
        }
        else if (statusFilter == "Healthy")
        {
            notebooks = notebooks.Where(n => !n.IsBroken).ToList();
        }
        // "All" → no filter

        // Apply search filter
        var searchText = _searchTextBox.Text.Trim();
        if (!string.IsNullOrEmpty(searchText))
        {
            notebooks = notebooks.Where(n =>
                n.FolderName.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                n.FolderServerRelativeUrl.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                n.LibraryTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        foreach (var notebook in notebooks)
        {
            var rowIndex = _notebooksGrid.Rows.Add(
                notebook.SiteUrl,
                notebook.LibraryTitle,
                notebook.FolderName,
                notebook.ItemId,
                notebook.StatusDescription,
                notebook.HtmlFileType
            );

            // Color coding
            if (notebook.IsFixed)
            {
                _notebooksGrid.Rows[rowIndex].DefaultCellStyle.BackColor = Color.FromArgb(200, 255, 200); // green
            }
            else if (notebook.IsBroken)
            {
                _notebooksGrid.Rows[rowIndex].DefaultCellStyle.BackColor = Color.FromArgb(255, 200, 200); // red
            }
            else
            {
                _notebooksGrid.Rows[rowIndex].DefaultCellStyle.BackColor = Color.FromArgb(200, 255, 200); // green
            }
        }

        SetStatus($"Showing {notebooks.Count} notebooks | Total: {result.TotalNotebooksFound} found, {result.TotalBroken} broken, {result.TotalFixed} fixed");
    }

    private void DisplaySiteSummary(BrokenOneNoteReportResult result)
    {
        _summaryList.Items.Clear();

        foreach (var site in result.SiteResults)
        {
            var item = new ListViewItem(site.SiteUrl);
            item.SubItems.Add(site.SiteTitle);
            item.SubItems.Add(site.TotalNotebooks.ToString());
            item.SubItems.Add(site.BrokenCount.ToString());
            item.SubItems.Add(site.FixedCount.ToString());
            item.SubItems.Add(site.Success ? "Success" : "Failed");
            item.SubItems.Add(site.ErrorMessage ?? "");

            if (!site.Success)
            {
                item.BackColor = Color.FromArgb(255, 200, 200);
            }
            else if (site.BrokenCount > 0)
            {
                item.BackColor = Color.FromArgb(255, 230, 200); // orange-ish for sites with broken notebooks
            }
            else if (site.TotalNotebooks == 0)
            {
                item.BackColor = Color.FromArgb(255, 255, 200); // yellow for no notebooks
            }
            else
            {
                item.BackColor = Color.FromArgb(200, 255, 200); // green for all healthy
            }

            _summaryList.Items.Add(item);
        }
    }

    private void FilterCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayNotebooks(_currentResult);
        }
    }

    private void SearchTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayNotebooks(_currentResult);
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
        _runButton.Text = "Cancel";
        _fixButton.Enabled = false;
        _deleteButton.Enabled = false;
        _exportButton.Enabled = false;
        _exportSummaryButton.Enabled = false;

        var progressPanel = Controls.Find("ProgressPanel", false).FirstOrDefault();
        if (progressPanel != null)
        {
            progressPanel.Visible = true;
        }
        _progressBar.Style = ProgressBarStyle.Marquee;
        _progressBar.MarqueeAnimationSpeed = 30;
        _progressLabel.Text = "Starting...";

        _cancellationTokenSource = new CancellationTokenSource();
        _notebooksGrid.Rows.Clear();
        _summaryList.Items.Clear();
        _logTextBox.Clear();
        _filterSiteCombo.Items.Clear();
        _filterSiteCombo.Items.Add("All Sites");
        _filterSiteCombo.SelectedIndex = 0;
        _filterStatusCombo.SelectedIndex = 0;

        var progress = new Progress<TaskProgress>(p =>
        {
            _progressLabel.Text = p.Message;

            if (_currentResult != null && _currentResult.ExecutionLog.Count > 0)
            {
                _logTextBox.Text = string.Join(Environment.NewLine, _currentResult.ExecutionLog);
                _logTextBox.SelectionStart = _logTextBox.TextLength;
                _logTextBox.ScrollToCaret();
            }
        });

        try
        {
            _currentResult = await _taskService.ExecuteBrokenOneNoteReportAsync(
                _task,
                _authService,
                progress,
                _cancellationTokenSource.Token);

            DisplayResults(_currentResult);

            if (_currentResult.Success)
            {
                SetStatus($"Task completed. {_currentResult.TotalNotebooksFound} notebooks found, {_currentResult.TotalBroken} broken across {_currentResult.SuccessfulSites} sites.");
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
            _runButton.Text = "\u25B6 Run Task";
            _deleteButton.Enabled = true;
            _exportButton.Enabled = _currentResult != null;
            _exportSummaryButton.Enabled = _currentResult != null;
            _fixButton.Enabled = _currentResult != null && _currentResult.TotalBroken > 0;
            _progressBar.Style = ProgressBarStyle.Blocks;
            if (progressPanel != null)
            {
                progressPanel.Visible = false;
            }
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;

            var updatedTask = await _taskService.GetTaskAsync(_task.Id);
            if (updatedTask != null)
            {
                _task = updatedTask;
                _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Status: {_task.StatusDescription}";
            }
        }
    }

    private async void FixButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null) return;

        var brokenCount = _currentResult.TotalBroken;
        if (brokenCount == 0)
        {
            MessageBox.Show("No broken notebooks to fix.", "Info",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var confirmResult = MessageBox.Show(
            $"This will set HTML_x0020_File_x0020_Type = \"OneNote.Notebook\" on {brokenCount} broken notebook folder(s).\n\n" +
            "This change is non-destructive and can be reversed.\n\nProceed?",
            "Fix Broken OneNote Notebooks",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (confirmResult != DialogResult.Yes) return;

        _runButton.Enabled = false;
        _fixButton.Enabled = false;
        _deleteButton.Enabled = false;

        var progressPanel = Controls.Find("ProgressPanel", false).FirstOrDefault();
        if (progressPanel != null)
        {
            progressPanel.Visible = true;
        }
        _progressBar.Style = ProgressBarStyle.Marquee;
        _progressBar.MarqueeAnimationSpeed = 30;
        _progressLabel.Text = "Fixing broken notebooks...";

        _cancellationTokenSource = new CancellationTokenSource();

        var progress = new Progress<TaskProgress>(p =>
        {
            _progressLabel.Text = p.Message;

            // Refresh display as fixes are applied
            if (_currentResult != null)
            {
                DisplayNotebooks(_currentResult);
                _logTextBox.Text = string.Join(Environment.NewLine, _currentResult.ExecutionLog);
                _logTextBox.SelectionStart = _logTextBox.TextLength;
                _logTextBox.ScrollToCaret();
            }
        });

        try
        {
            _currentResult = await _taskService.ExecuteBrokenOneNoteFixAsync(
                _task,
                _authService,
                _currentResult,
                progress,
                _cancellationTokenSource.Token);

            DisplayResults(_currentResult);
            SetStatus($"Fix completed. {_currentResult.TotalFixed} fixed, {_currentResult.TotalBroken} still broken.");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fix failed: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Fix failed");
        }
        finally
        {
            _runButton.Enabled = true;
            _fixButton.Enabled = _currentResult != null && _currentResult.TotalBroken > 0;
            _deleteButton.Enabled = true;
            _progressBar.Style = ProgressBarStyle.Blocks;
            if (progressPanel != null)
            {
                progressPanel.Visible = false;
            }
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }
    }

    private void ExportButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var brokenOnly = _filterStatusCombo.SelectedItem?.ToString() == "Broken Only";

        var safeName = SanitizeFileName(_task.Name);
        var suffix = brokenOnly ? "Broken" : "AllNotebooks";
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_{suffix}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportBrokenOneNoteReport(_currentResult, dialog.FileName, brokenOnly);
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

    private void ExportSummaryButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_Summary_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportBrokenOneNoteReportSummary(_currentResult, dialog.FileName);
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

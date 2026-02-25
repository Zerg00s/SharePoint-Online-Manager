using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing customized lists report task details and results.
/// </summary>
public class CustomizedListsDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _exportButton = null!;
    private Button _exportSummaryButton = null!;
    private Button _deleteButton = null!;
    private DataGridView _listsGrid = null!;
    private ListView _summaryList = null!;
    private TextBox _logTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private ComboBox _filterSiteCombo = null!;
    private ComboBox _filterFormTypeCombo = null!;
    private TextBox _searchTextBox = null!;

    private TaskDefinition _task = null!;
    private CustomizedListsReportResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private CsvExporter _csvExporter = null!;

    public override string ScreenTitle => _task?.Name ?? "Customized Lists Report";

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
            Size = new Size(800, 35),
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

        buttonPanel.Controls.AddRange(new Control[] { _runButton, _exportButton, _exportSummaryButton, _deleteButton });

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

        // Customized Lists tab
        var listsTab = new TabPage("Lists");

        var listsHeaderPanel = new FlowLayoutPanel
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

        var filterFormTypeLabel = new Label
        {
            Text = "Form Type:",
            AutoSize = true,
            Padding = new Padding(10, 5, 5, 0)
        };

        _filterFormTypeCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 140
        };
        _filterFormTypeCombo.Items.Add("Customized Only");
        _filterFormTypeCombo.Items.Add("All Lists");
        _filterFormTypeCombo.Items.Add("Power Apps");
        _filterFormTypeCombo.Items.Add("SPFx Custom Form");
        _filterFormTypeCombo.Items.Add("Default");
        _filterFormTypeCombo.SelectedIndex = 0;
        _filterFormTypeCombo.SelectedIndexChanged += FilterCombo_SelectedIndexChanged;

        var searchLabel = new Label
        {
            Text = "Search:",
            AutoSize = true,
            Padding = new Padding(10, 5, 5, 0)
        };

        _searchTextBox = new TextBox
        {
            Width = 150,
            PlaceholderText = "List name..."
        };
        _searchTextBox.TextChanged += SearchTextBox_TextChanged;

        listsHeaderPanel.Controls.AddRange(new Control[]
        {
            filterSiteLabel, _filterSiteCombo,
            filterFormTypeLabel, _filterFormTypeCombo,
            searchLabel, _searchTextBox
        });

        _listsGrid = new DataGridView
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
            if (_listsGrid.CurrentCell != null)
            {
                Clipboard.SetText(_listsGrid.CurrentCell.Value?.ToString() ?? "");
            }
        });
        copyMenuItem.ShortcutKeys = Keys.Control | Keys.C;
        var copyRowMenuItem = new ToolStripMenuItem("Copy Row", null, (s, e) =>
        {
            if (_listsGrid.CurrentRow != null)
            {
                var values = new List<string>();
                foreach (DataGridViewCell cell in _listsGrid.CurrentRow.Cells)
                {
                    values.Add(cell.Value?.ToString() ?? "");
                }
                Clipboard.SetText(string.Join("\t", values));
            }
        });
        gridContextMenu.Items.Add(copyMenuItem);
        gridContextMenu.Items.Add(copyRowMenuItem);
        _listsGrid.ContextMenuStrip = gridContextMenu;

        _listsGrid.Columns.Add("SiteUrl", "Site URL");
        _listsGrid.Columns.Add("SiteTitle", "Site Title");
        _listsGrid.Columns.Add("ListTitle", "List Name");
        _listsGrid.Columns.Add("ListType", "List Type");
        _listsGrid.Columns.Add("FormType", "Form Type");
        _listsGrid.Columns.Add("ItemCount", "Items");
        _listsGrid.Columns.Add("ListUrl", "List URL");

        _listsGrid.Columns["SiteUrl"].FillWeight = 140;
        _listsGrid.Columns["SiteTitle"].FillWeight = 100;
        _listsGrid.Columns["ListTitle"].FillWeight = 120;
        _listsGrid.Columns["ListType"].FillWeight = 80;
        _listsGrid.Columns["FormType"].FillWeight = 90;
        _listsGrid.Columns["ItemCount"].FillWeight = 50;
        _listsGrid.Columns["ListUrl"].FillWeight = 140;

        listsTab.Controls.Add(_listsGrid);
        listsTab.Controls.Add(listsHeaderPanel);

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
        _summaryList.Columns.Add("Total Lists", 80);
        _summaryList.Columns.Add("Customized", 80);
        _summaryList.Columns.Add("Power Apps", 80);
        _summaryList.Columns.Add("SPFx", 60);
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

        _tabControl.TabPages.AddRange(new[] { listsTab, summaryTab, logTab });

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

        _currentResult = await _taskService.GetLatestCustomizedListsReportResultAsync(_task.Id);

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportButton.Enabled = true;
            _exportSummaryButton.Enabled = true;
        }
        else
        {
            _listsGrid.Rows.Clear();
            _summaryList.Items.Clear();
            _logTextBox.Text = "No results yet. Click 'Run Task' to execute.";
            _exportButton.Enabled = false;
            _exportSummaryButton.Enabled = false;
        }
    }

    private void DisplayResults(CustomizedListsReportResult result)
    {
        PopulateSiteFilter(result);
        DisplayLists(result);
        DisplaySiteSummary(result);
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);
    }

    private void PopulateSiteFilter(CustomizedListsReportResult result)
    {
        _filterSiteCombo.Items.Clear();
        _filterSiteCombo.Items.Add("All Sites");

        foreach (var site in result.SiteResults.Where(s => s.Success))
        {
            _filterSiteCombo.Items.Add(site.SiteUrl);
        }

        _filterSiteCombo.SelectedIndex = 0;
    }

    private void DisplayLists(CustomizedListsReportResult result)
    {
        _listsGrid.Rows.Clear();

        var lists = result.GetAllLists().ToList();

        // Apply site filter
        var selectedSite = _filterSiteCombo.SelectedIndex > 0 ? _filterSiteCombo.SelectedItem?.ToString() : null;
        if (!string.IsNullOrEmpty(selectedSite))
        {
            lists = lists.Where(l => l.SiteUrl.Equals(selectedSite, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        // Apply form type filter
        var formTypeFilter = _filterFormTypeCombo.SelectedItem?.ToString();
        if (formTypeFilter == "Customized Only")
        {
            lists = lists.Where(l => l.IsCustomized).ToList();
        }
        else if (formTypeFilter == "Power Apps")
        {
            lists = lists.Where(l => l.FormType == ListFormType.PowerApps).ToList();
        }
        else if (formTypeFilter == "SPFx Custom Form")
        {
            lists = lists.Where(l => l.FormType == ListFormType.SPFxCustomForm).ToList();
        }
        else if (formTypeFilter == "Default")
        {
            lists = lists.Where(l => l.FormType == ListFormType.Default).ToList();
        }
        // "All Lists" → no filter

        // Apply search filter
        var searchText = _searchTextBox.Text.Trim();
        if (!string.IsNullOrEmpty(searchText))
        {
            lists = lists.Where(l =>
                l.ListTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                l.ListUrl.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        foreach (var list in lists)
        {
            var rowIndex = _listsGrid.Rows.Add(
                list.SiteUrl,
                list.SiteTitle,
                list.ListTitle,
                list.ListType,
                list.FormTypeDescription,
                list.ItemCount,
                list.ListUrl
            );

            // Color code customized rows
            if (list.FormType == ListFormType.PowerApps)
            {
                _listsGrid.Rows[rowIndex].DefaultCellStyle.BackColor = Color.FromArgb(230, 210, 255); // light purple
            }
            else if (list.FormType == ListFormType.SPFxCustomForm)
            {
                _listsGrid.Rows[rowIndex].DefaultCellStyle.BackColor = Color.FromArgb(210, 230, 255); // light blue
            }
        }

        SetStatus($"Showing {lists.Count} lists | Total: {result.TotalListsScanned} scanned, {result.TotalCustomized} customized ({result.TotalPowerApps} Power Apps, {result.TotalSpfx} SPFx)");
    }

    private void DisplaySiteSummary(CustomizedListsReportResult result)
    {
        _summaryList.Items.Clear();

        foreach (var site in result.SiteResults)
        {
            var item = new ListViewItem(site.SiteUrl);
            item.SubItems.Add(site.SiteTitle);
            item.SubItems.Add(site.TotalLists.ToString());
            item.SubItems.Add(site.CustomizedCount.ToString());
            item.SubItems.Add(site.PowerAppsCount.ToString());
            item.SubItems.Add(site.SpfxCount.ToString());
            item.SubItems.Add(site.Success ? "Success" : "Failed");
            item.SubItems.Add(site.ErrorMessage ?? "");

            if (!site.Success)
            {
                item.BackColor = Color.FromArgb(255, 200, 200);
            }
            else if (site.CustomizedCount == 0)
            {
                item.BackColor = Color.FromArgb(255, 255, 200);
            }
            else
            {
                item.BackColor = Color.FromArgb(200, 255, 200);
            }

            _summaryList.Items.Add(item);
        }
    }

    private void FilterCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayLists(_currentResult);
        }
    }

    private void SearchTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayLists(_currentResult);
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
        _listsGrid.Rows.Clear();
        _summaryList.Items.Clear();
        _logTextBox.Clear();
        _filterSiteCombo.Items.Clear();
        _filterSiteCombo.Items.Add("All Sites");
        _filterSiteCombo.SelectedIndex = 0;
        _filterFormTypeCombo.SelectedIndex = 0;

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
            _currentResult = await _taskService.ExecuteCustomizedListsReportAsync(
                _task,
                _authService,
                progress,
                _cancellationTokenSource.Token);

            DisplayResults(_currentResult);

            if (_currentResult.Success)
            {
                SetStatus($"Task completed. {_currentResult.TotalCustomized} customized lists found ({_currentResult.TotalPowerApps} Power Apps, {_currentResult.TotalSpfx} SPFx) across {_currentResult.SuccessfulSites} sites.");
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

    private void ExportButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var customizedOnly = _filterFormTypeCombo.SelectedItem?.ToString() != "All Lists" &&
                             _filterFormTypeCombo.SelectedItem?.ToString() != "Default";

        var safeName = SanitizeFileName(_task.Name);
        var suffix = customizedOnly ? "Customized" : "AllLists";
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_{suffix}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportCustomizedListsReport(_currentResult, dialog.FileName, customizedOnly);
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
                _csvExporter.ExportCustomizedListsReportSummary(_currentResult, dialog.FileName);
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

using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing site access check task details and results.
/// </summary>
public class SiteAccessDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _exportAllButton = null!;
    private Button _exportSourceIssuesButton = null!;
    private Button _exportTargetIssuesButton = null!;
    private Button _deleteButton = null!;
    private Button _createSourceAdminTaskButton = null!;
    private Button _createTargetAdminTaskButton = null!;
    private ListView _summaryListView = null!;
    private ListView _sourceIssuesListView = null!;
    private ListView _targetIssuesListView = null!;
    private TextBox _logTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;

    private TaskDefinition _task = null!;
    private SiteAccessResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private CsvExporter _csvExporter = null!;

    public override string ScreenTitle => _task?.Name ?? "Site Access Check Details";

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
            Size = new Size(1000, 35),
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

        _exportSourceIssuesButton = new Button
        {
            Text = "\u26A0 Source Issues",
            Size = new Size(120, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportSourceIssuesButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportSourceIssuesButton.FlatAppearance.BorderSize = 1;
        _exportSourceIssuesButton.Click += ExportSourceIssuesButton_Click;

        _exportTargetIssuesButton = new Button
        {
            Text = "\u26A0 Target Issues",
            Size = new Size(120, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportTargetIssuesButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportTargetIssuesButton.FlatAppearance.BorderSize = 1;
        _exportTargetIssuesButton.Click += ExportTargetIssuesButton_Click;

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

        buttonPanel.Controls.AddRange(new Control[] { _runButton, _exportAllButton, _exportSourceIssuesButton, _exportTargetIssuesButton, _deleteButton });

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

        // Tab control for results
        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        // Summary tab
        var summaryTab = new TabPage("Summary");
        _summaryListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _summaryListView.Columns.Add("Source Site", 300);
        _summaryListView.Columns.Add("Source Status", 100);
        _summaryListView.Columns.Add("Target Site", 300);
        _summaryListView.Columns.Add("Target Status", 100);
        _summaryListView.ContextMenuStrip = CreateListViewContextMenu(_summaryListView);
        summaryTab.Controls.Add(_summaryListView);

        // Source Issues tab
        var sourceIssuesTab = new TabPage("Source Issues");

        var sourceIssuesButtonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 40,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(5)
        };

        _createSourceAdminTaskButton = new Button
        {
            Text = "\U0001F464 Create Add Admin Task",
            Size = new Size(180, 28),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _createSourceAdminTaskButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _createSourceAdminTaskButton.FlatAppearance.BorderSize = 1;
        _createSourceAdminTaskButton.Click += CreateSourceAdminTaskButton_Click;
        sourceIssuesButtonPanel.Controls.Add(_createSourceAdminTaskButton);

        _sourceIssuesListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _sourceIssuesListView.Columns.Add("Site URL", 400);
        _sourceIssuesListView.Columns.Add("Site Title", 150);
        _sourceIssuesListView.Columns.Add("Status", 100);
        _sourceIssuesListView.Columns.Add("Account", 150);
        _sourceIssuesListView.Columns.Add("Error", 300);
        _sourceIssuesListView.ContextMenuStrip = CreateListViewContextMenu(_sourceIssuesListView);

        sourceIssuesTab.Controls.Add(_sourceIssuesListView);
        sourceIssuesTab.Controls.Add(sourceIssuesButtonPanel);

        // Target Issues tab
        var targetIssuesTab = new TabPage("Target Issues");

        var targetIssuesButtonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 40,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(5)
        };

        _createTargetAdminTaskButton = new Button
        {
            Text = "\U0001F464 Create Add Admin Task",
            Size = new Size(180, 28),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _createTargetAdminTaskButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _createTargetAdminTaskButton.FlatAppearance.BorderSize = 1;
        _createTargetAdminTaskButton.Click += CreateTargetAdminTaskButton_Click;
        targetIssuesButtonPanel.Controls.Add(_createTargetAdminTaskButton);

        _targetIssuesListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _targetIssuesListView.Columns.Add("Site URL", 400);
        _targetIssuesListView.Columns.Add("Site Title", 150);
        _targetIssuesListView.Columns.Add("Status", 100);
        _targetIssuesListView.Columns.Add("Account", 150);
        _targetIssuesListView.Columns.Add("Error", 300);
        _targetIssuesListView.ContextMenuStrip = CreateListViewContextMenu(_targetIssuesListView);

        targetIssuesTab.Controls.Add(_targetIssuesListView);
        targetIssuesTab.Controls.Add(targetIssuesButtonPanel);

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

        _tabControl.TabPages.AddRange(new[] { summaryTab, sourceIssuesTab, targetIssuesTab, logTab });

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
        _currentResult = await _taskService.GetLatestSiteAccessResultAsync(_task.Id);

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportAllButton.Enabled = true;
            _exportSourceIssuesButton.Enabled = _currentResult.GetSourceIssues().Any();
            _exportTargetIssuesButton.Enabled = _currentResult.GetTargetIssues().Any();
            _createSourceAdminTaskButton.Enabled = _currentResult.GetSourceIssues().Any();
            _createTargetAdminTaskButton.Enabled = _currentResult.GetTargetIssues().Any();
        }
        else
        {
            _summaryListView.Items.Clear();
            _sourceIssuesListView.Items.Clear();
            _targetIssuesListView.Items.Clear();
            _logTextBox.Text = "No results yet. Click 'Run Task' to execute.";
            _exportAllButton.Enabled = false;
            _exportSourceIssuesButton.Enabled = false;
            _exportTargetIssuesButton.Enabled = false;
            _createSourceAdminTaskButton.Enabled = false;
            _createTargetAdminTaskButton.Enabled = false;
        }
    }

    private void DisplayResults(SiteAccessResult result)
    {
        DisplaySummary(result);
        DisplaySourceIssues(result);
        DisplayTargetIssues(result);
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);

        // Update tab titles with counts
        _tabControl.TabPages[1].Text = $"Source Issues ({result.GetSourceIssues().Count()})";
        _tabControl.TabPages[2].Text = $"Target Issues ({result.GetTargetIssues().Count()})";
    }

    private void DisplaySummary(SiteAccessResult result)
    {
        _summaryListView.Items.Clear();

        foreach (var pair in result.PairResults)
        {
            var item = new ListViewItem(pair.SourceSiteUrl);
            item.SubItems.Add(pair.SourceResult.StatusDescription);
            item.SubItems.Add(pair.TargetSiteUrl);
            item.SubItems.Add(pair.TargetResult.StatusDescription);

            // Color code based on issues
            if (pair.HasSourceIssue && pair.HasTargetIssue)
            {
                item.BackColor = Color.FromArgb(255, 200, 200); // Red for both issues
            }
            else if (pair.HasSourceIssue || pair.HasTargetIssue)
            {
                item.BackColor = Color.FromArgb(255, 255, 200); // Yellow for one issue
            }
            else
            {
                item.BackColor = Color.FromArgb(200, 255, 200); // Green for accessible
            }

            _summaryListView.Items.Add(item);
        }
    }

    private void DisplaySourceIssues(SiteAccessResult result)
    {
        _sourceIssuesListView.Items.Clear();

        foreach (var issue in result.GetSourceIssues())
        {
            var item = new ListViewItem(issue.SiteUrl);
            item.SubItems.Add(issue.SiteTitle);
            item.SubItems.Add(issue.StatusDescription);
            item.SubItems.Add(issue.AccountUsed ?? "");
            item.SubItems.Add(issue.ErrorMessage ?? "");

            // Color code by status
            item.BackColor = issue.Status switch
            {
                SiteAccessStatus.AccessDenied => Color.FromArgb(255, 220, 220),
                SiteAccessStatus.NotFound => Color.FromArgb(255, 240, 200),
                SiteAccessStatus.AuthenticationRequired => Color.FromArgb(255, 200, 255),
                _ => Color.FromArgb(220, 220, 220)
            };

            _sourceIssuesListView.Items.Add(item);
        }
    }

    private void DisplayTargetIssues(SiteAccessResult result)
    {
        _targetIssuesListView.Items.Clear();

        foreach (var issue in result.GetTargetIssues())
        {
            var item = new ListViewItem(issue.SiteUrl);
            item.SubItems.Add(issue.SiteTitle);
            item.SubItems.Add(issue.StatusDescription);
            item.SubItems.Add(issue.AccountUsed ?? "");
            item.SubItems.Add(issue.ErrorMessage ?? "");

            // Color code by status
            item.BackColor = issue.Status switch
            {
                SiteAccessStatus.AccessDenied => Color.FromArgb(255, 220, 220),
                SiteAccessStatus.NotFound => Color.FromArgb(255, 240, 200),
                SiteAccessStatus.AuthenticationRequired => Color.FromArgb(255, 200, 255),
                _ => Color.FromArgb(220, 220, 220)
            };

            _targetIssuesListView.Items.Add(item);
        }
    }

    private void AddPairToSummaryList(SitePairAccessResult pair)
    {
        var item = new ListViewItem(pair.SourceSiteUrl);
        item.SubItems.Add(pair.SourceResult.StatusDescription);
        item.SubItems.Add(pair.TargetSiteUrl);
        item.SubItems.Add(pair.TargetResult.StatusDescription);

        // Color code based on issues
        if (pair.HasSourceIssue && pair.HasTargetIssue)
        {
            item.BackColor = Color.FromArgb(255, 200, 200); // Red for both issues
        }
        else if (pair.HasSourceIssue || pair.HasTargetIssue)
        {
            item.BackColor = Color.FromArgb(255, 255, 200); // Yellow for one issue
        }
        else
        {
            item.BackColor = Color.FromArgb(200, 255, 200); // Green for accessible
        }

        _summaryListView.Items.Add(item);
    }

    private void AddToSourceIssuesList(SiteAccessCheckItem issue)
    {
        var item = new ListViewItem(issue.SiteUrl);
        item.SubItems.Add(issue.SiteTitle);
        item.SubItems.Add(issue.StatusDescription);
        item.SubItems.Add(issue.AccountUsed ?? "");
        item.SubItems.Add(issue.ErrorMessage ?? "");

        // Color code by status
        item.BackColor = issue.Status switch
        {
            SiteAccessStatus.AccessDenied => Color.FromArgb(255, 220, 220),
            SiteAccessStatus.NotFound => Color.FromArgb(255, 240, 200),
            SiteAccessStatus.AuthenticationRequired => Color.FromArgb(255, 200, 255),
            _ => Color.FromArgb(220, 220, 220)
        };

        _sourceIssuesListView.Items.Add(item);
    }

    private void AddToTargetIssuesList(SiteAccessCheckItem issue)
    {
        var item = new ListViewItem(issue.SiteUrl);
        item.SubItems.Add(issue.SiteTitle);
        item.SubItems.Add(issue.StatusDescription);
        item.SubItems.Add(issue.AccountUsed ?? "");
        item.SubItems.Add(issue.ErrorMessage ?? "");

        // Color code by status
        item.BackColor = issue.Status switch
        {
            SiteAccessStatus.AccessDenied => Color.FromArgb(255, 220, 220),
            SiteAccessStatus.NotFound => Color.FromArgb(255, 240, 200),
            SiteAccessStatus.AuthenticationRequired => Color.FromArgb(255, 200, 255),
            _ => Color.FromArgb(220, 220, 220)
        };

        _targetIssuesListView.Items.Add(item);
    }

    private async void RunButton_Click(object? sender, EventArgs e)
    {
        // If task is running, cancel it
        if (_cancellationTokenSource != null)
        {
            _cancellationTokenSource.Cancel();
            return;
        }

        await ExecuteTaskAsync();
    }

    private async Task ExecuteTaskAsync()
    {
        var progressPanel = Controls.Find("ProgressPanel", true).FirstOrDefault();
        if (progressPanel != null)
        {
            progressPanel.Visible = true;
        }

        _runButton.Text = "\u23F9 Cancel";
        _runButton.BackColor = Color.FromArgb(200, 50, 50);
        _runButton.FlatAppearance.BorderColor = Color.FromArgb(200, 50, 50);
        _exportAllButton.Enabled = false;
        _exportSourceIssuesButton.Enabled = false;
        _exportTargetIssuesButton.Enabled = false;

        _cancellationTokenSource = new CancellationTokenSource();

        // Clear previous results
        _summaryListView.Items.Clear();
        _sourceIssuesListView.Items.Clear();
        _targetIssuesListView.Items.Clear();
        _logTextBox.Clear();

        try
        {
            var progress = new Progress<TaskProgress>(p =>
            {
                _progressBar.Value = p.PercentComplete;
                _progressLabel.Text = p.Message;

                // Real-time update: add completed pair to UI
                if (p.CompletedAccessPairResult != null)
                {
                    AddPairToSummaryList(p.CompletedAccessPairResult);

                    if (p.CompletedAccessPairResult.HasSourceIssue)
                    {
                        AddToSourceIssuesList(p.CompletedAccessPairResult.SourceResult);
                    }

                    if (p.CompletedAccessPairResult.HasTargetIssue)
                    {
                        AddToTargetIssuesList(p.CompletedAccessPairResult.TargetResult);
                    }

                    // Update tab titles with counts
                    _tabControl.TabPages[1].Text = $"Source Issues ({_sourceIssuesListView.Items.Count})";
                    _tabControl.TabPages[2].Text = $"Target Issues ({_targetIssuesListView.Items.Count})";
                }
            });

            _currentResult = await _taskService.ExecuteSiteAccessCheckAsync(
                _task,
                _authService,
                _connectionManager,
                progress,
                _cancellationTokenSource.Token);

            await RefreshTaskDetailsAsync();

            if (_currentResult.Success)
            {
                var sourceIssues = _currentResult.GetSourceIssues().Count();
                var targetIssues = _currentResult.GetTargetIssues().Count();

                SetStatus($"Task completed. Source issues: {sourceIssues}, Target issues: {targetIssues}");
            }
            else
            {
                SetStatus($"Task failed: {_currentResult.ErrorMessage}");
            }
        }
        catch (OperationCanceledException)
        {
            SetStatus("Task cancelled by user");
            await RefreshTaskDetailsAsync();
        }
        catch (Exception ex)
        {
            SetStatus($"Error: {ex.Message}");
            MessageBox.Show($"Task execution failed: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (progressPanel != null)
            {
                progressPanel.Visible = false;
            }

            _runButton.Text = "\u25B6 Run Task";
            _runButton.BackColor = Color.FromArgb(0, 120, 212);
            _runButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }
    }

    private void ExportAllButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null) return;

        var reportsFolder = GetReportsFolder();

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            Title = "Export Site Access Report",
            InitialDirectory = reportsFolder,
            FileName = $"SiteAccess_{_task.Name}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK) return;

        try
        {
            _csvExporter.ExportSiteAccessReport(_currentResult, dialog.FileName);
            SetStatus($"Exported to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Export Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ExportSourceIssuesButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null) return;

        var reportsFolder = GetReportsFolder();

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            Title = "Export Source Access Issues",
            InitialDirectory = reportsFolder,
            FileName = $"SourceAccessIssues_{_task.Name}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK) return;

        try
        {
            _csvExporter.ExportSiteAccessIssues(_currentResult, dialog.FileName, sourceIssuesOnly: true);
            SetStatus($"Exported to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Export Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ExportTargetIssuesButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null) return;

        var reportsFolder = GetReportsFolder();

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            Title = "Export Target Access Issues",
            InitialDirectory = reportsFolder,
            FileName = $"TargetAccessIssues_{_task.Name}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() != DialogResult.OK) return;

        try
        {
            _csvExporter.ExportSiteAccessIssues(_currentResult, dialog.FileName, targetIssuesOnly: true);
            SetStatus($"Exported to {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed: {ex.Message}", "Export Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async void DeleteButton_Click(object? sender, EventArgs e)
    {
        var result = MessageBox.Show(
            $"Are you sure you want to delete the task '{_task.Name}' and all its results?",
            "Delete Task",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (result != DialogResult.Yes) return;

        await _taskService.DeleteTaskAsync(_task.Id);
        SetStatus("Task deleted");
        await NavigationService!.GoBackAsync();
    }

    private static string GetReportsFolder()
    {
        var folder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SharePointOnlineManager",
            "Reports");
        Directory.CreateDirectory(folder);
        return folder;
    }

    private static ContextMenuStrip CreateListViewContextMenu(ListView listView)
    {
        var contextMenu = new ContextMenuStrip();

        var copyCell = new ToolStripMenuItem("Copy Cell");
        copyCell.Click += (s, e) =>
        {
            if (listView.SelectedItems.Count > 0)
            {
                var item = listView.SelectedItems[0];
                var point = listView.PointToClient(Cursor.Position);
                var hitTest = listView.HitTest(point);

                if (hitTest.SubItem != null)
                {
                    Clipboard.SetText(hitTest.SubItem.Text ?? "");
                }
                else if (hitTest.Item != null)
                {
                    Clipboard.SetText(hitTest.Item.Text ?? "");
                }
            }
        };

        var copyRow = new ToolStripMenuItem("Copy Row");
        copyRow.Click += (s, e) =>
        {
            if (listView.SelectedItems.Count > 0)
            {
                var item = listView.SelectedItems[0];
                var values = new List<string> { item.Text };
                foreach (ListViewItem.ListViewSubItem subItem in item.SubItems)
                {
                    if (subItem != item.SubItems[0]) // Skip first since it's already added
                        values.Add(subItem.Text);
                }
                Clipboard.SetText(string.Join("\t", values));
            }
        };

        var copyAllUrls = new ToolStripMenuItem("Copy All URLs");
        copyAllUrls.Click += (s, e) =>
        {
            var urls = new List<string>();
            foreach (ListViewItem item in listView.Items)
            {
                urls.Add(item.Text);
            }
            if (urls.Count > 0)
            {
                Clipboard.SetText(string.Join(Environment.NewLine, urls));
            }
        };

        contextMenu.Items.Add(copyCell);
        contextMenu.Items.Add(copyRow);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(copyAllUrls);

        return contextMenu;
    }

    private async void CreateSourceAdminTaskButton_Click(object? sender, EventArgs e)
    {
        await CreateAddAdminTaskFromIssuesAsync(isSource: true);
    }

    private async void CreateTargetAdminTaskButton_Click(object? sender, EventArgs e)
    {
        await CreateAddAdminTaskFromIssuesAsync(isSource: false);
    }

    private async Task CreateAddAdminTaskFromIssuesAsync(bool isSource)
    {
        if (_currentResult == null) return;

        // Get the issues for the appropriate side
        var issues = isSource
            ? _currentResult.GetSourceIssues().ToList()
            : _currentResult.GetTargetIssues().ToList();

        if (issues.Count == 0)
        {
            MessageBox.Show("No issues found to create task from.", "No Issues",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Get the connection ID from the task configuration
        var config = System.Text.Json.JsonSerializer.Deserialize<SiteAccessConfiguration>(
            _task.ConfigurationJson ?? "{}",
            new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        if (config == null)
        {
            MessageBox.Show("Could not read task configuration.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var connectionId = isSource ? config.SourceConnectionId : config.TargetConnectionId;
        var connection = await _connectionManager.GetConnectionAsync(connectionId);

        if (connection == null)
        {
            MessageBox.Show("Connection not found.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var siteUrls = issues.Select(i => i.SiteUrl).ToList();
        var sideName = isSource ? "Source" : "Target";

        var result = MessageBox.Show(
            $"Create an 'Add Site Collection Admins' task for {siteUrls.Count} {sideName} site(s) with access issues?\n\n" +
            $"Connection: {connection.Name}",
            "Create Task",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result != DialogResult.Yes) return;

        // Create the context and navigate to the config screen
        var context = new TaskCreationContext
        {
            Connection = connection,
            SelectedSites = siteUrls.Select(url => new SiteCollection { Url = url }).ToList()
        };

        await NavigationService!.NavigateToAsync<AddSiteAdminsConfigScreen>(context);
    }
}

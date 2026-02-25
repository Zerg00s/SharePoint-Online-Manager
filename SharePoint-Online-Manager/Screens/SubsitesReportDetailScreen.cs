using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing subsites report task details and results.
/// </summary>
public class SubsitesReportDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _exportSubsitesButton = null!;
    private Button _exportSummaryButton = null!;
    private Button _deleteButton = null!;
    private DataGridView _subsitesGrid = null!;
    private ListView _summaryList = null!;
    private TextBox _logTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private ComboBox _filterSiteCombo = null!;
    private TextBox _searchTextBox = null!;

    private TaskDefinition _task = null!;
    private SubsitesReportResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private CsvExporter _csvExporter = null!;

    public override string ScreenTitle => _task?.Name ?? "Subsites Report";

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

        _exportSubsitesButton = new Button
        {
            Text = "\U0001F4BE Export Subsites",
            Size = new Size(140, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportSubsitesButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportSubsitesButton.FlatAppearance.BorderSize = 1;
        _exportSubsitesButton.Click += ExportSubsitesButton_Click;

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

        buttonPanel.Controls.AddRange(new Control[] { _runButton, _exportSubsitesButton, _exportSummaryButton, _deleteButton });

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

        // Subsites tab
        var subsitesTab = new TabPage("Subsites");

        var subsitesHeaderPanel = new FlowLayoutPanel
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
            Width = 300
        };
        _filterSiteCombo.Items.Add("All Sites");
        _filterSiteCombo.Items.Add("Sites with Subsites Only");
        _filterSiteCombo.SelectedIndex = 0;
        _filterSiteCombo.SelectedIndexChanged += FilterCombo_SelectedIndexChanged;

        var searchLabel = new Label
        {
            Text = "Search:",
            AutoSize = true,
            Padding = new Padding(10, 5, 5, 0)
        };

        _searchTextBox = new TextBox
        {
            Width = 200,
            PlaceholderText = "URL or title..."
        };
        _searchTextBox.TextChanged += SearchTextBox_TextChanged;

        subsitesHeaderPanel.Controls.AddRange(new Control[]
        {
            filterSiteLabel, _filterSiteCombo,
            searchLabel, _searchTextBox
        });

        _subsitesGrid = new DataGridView
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
            if (_subsitesGrid.CurrentCell != null)
            {
                Clipboard.SetText(_subsitesGrid.CurrentCell.Value?.ToString() ?? "");
            }
        });
        copyMenuItem.ShortcutKeys = Keys.Control | Keys.C;
        var copyRowMenuItem = new ToolStripMenuItem("Copy Row", null, (s, e) =>
        {
            if (_subsitesGrid.CurrentRow != null)
            {
                var values = new List<string>();
                foreach (DataGridViewCell cell in _subsitesGrid.CurrentRow.Cells)
                {
                    values.Add(cell.Value?.ToString() ?? "");
                }
                Clipboard.SetText(string.Join("\t", values));
            }
        });
        gridContextMenu.Items.Add(copyMenuItem);
        gridContextMenu.Items.Add(copyRowMenuItem);
        _subsitesGrid.ContextMenuStrip = gridContextMenu;

        _subsitesGrid.Columns.Add("SiteCollectionUrl", "Site Collection URL");
        _subsitesGrid.Columns.Add("SiteUrl", "Parent Site URL");
        _subsitesGrid.Columns.Add("SiteTitle", "Parent Site Title");
        _subsitesGrid.Columns.Add("SubsiteUrl", "Subsite URL");
        _subsitesGrid.Columns.Add("SubsiteTitle", "Subsite Title");
        _subsitesGrid.Columns.Add("WebTemplate", "Template");
        _subsitesGrid.Columns.Add("Created", "Created");
        _subsitesGrid.Columns.Add("LastModified", "Last Modified");
        _subsitesGrid.Columns.Add("Language", "Language");

        _subsitesGrid.Columns["SiteCollectionUrl"]!.FillWeight = 140;
        _subsitesGrid.Columns["SiteUrl"]!.FillWeight = 140;
        _subsitesGrid.Columns["SiteTitle"]!.FillWeight = 100;
        _subsitesGrid.Columns["SubsiteUrl"]!.FillWeight = 160;
        _subsitesGrid.Columns["SubsiteTitle"]!.FillWeight = 100;
        _subsitesGrid.Columns["WebTemplate"]!.FillWeight = 70;
        _subsitesGrid.Columns["Created"]!.FillWeight = 80;
        _subsitesGrid.Columns["LastModified"]!.FillWeight = 80;
        _subsitesGrid.Columns["Language"]!.FillWeight = 70;

        subsitesTab.Controls.Add(_subsitesGrid);
        subsitesTab.Controls.Add(subsitesHeaderPanel);

        // Site Summary tab
        var summaryTab = new TabPage("Site Summary");

        _summaryList = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _summaryList.Columns.Add("Site URL", 300);
        _summaryList.Columns.Add("Site Title", 200);
        _summaryList.Columns.Add("Subsite Count", 100);
        _summaryList.Columns.Add("Status", 80);
        _summaryList.Columns.Add("Error", 300);

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

        _tabControl.TabPages.AddRange(new[] { subsitesTab, summaryTab, logTab });

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

        _currentResult = await _taskService.GetLatestSubsitesReportResultAsync(_task.Id);

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportSubsitesButton.Enabled = true;
            _exportSummaryButton.Enabled = true;
        }
        else
        {
            _subsitesGrid.Rows.Clear();
            _summaryList.Items.Clear();
            _logTextBox.Text = "No results yet. Click 'Run Task' to execute.";
            _exportSubsitesButton.Enabled = false;
            _exportSummaryButton.Enabled = false;
        }
    }

    private void DisplayResults(SubsitesReportResult result)
    {
        DisplaySubsites(result);
        DisplaySiteSummary(result);
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);
    }

    private void DisplaySubsites(SubsitesReportResult result)
    {
        _subsitesGrid.Rows.Clear();

        var subsites = result.GetAllSubsites().ToList();

        // Apply site filter
        var siteFilter = _filterSiteCombo.SelectedItem?.ToString();
        if (siteFilter == "Sites with Subsites Only")
        {
            var sitesWithSubsites = new HashSet<string>(
                result.SiteResults.Where(s => s.HasSubsites).Select(s => s.SiteUrl),
                StringComparer.OrdinalIgnoreCase);
            subsites = subsites.Where(s => sitesWithSubsites.Contains(s.SiteUrl)).ToList();
        }

        // If filterSiteCombo has specific sites selected (added dynamically)
        if (_filterSiteCombo.SelectedIndex > 1 && _filterSiteCombo.SelectedItem is string selectedSite)
        {
            subsites = subsites.Where(s =>
                s.SiteUrl.Equals(selectedSite, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        // Apply search filter
        var searchText = _searchTextBox.Text.Trim();
        if (!string.IsNullOrEmpty(searchText))
        {
            subsites = subsites.Where(s =>
                s.SubsiteUrl.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                s.SubsiteTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                s.SiteUrl.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                s.SiteTitle.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                s.WebTemplate.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        foreach (var subsite in subsites)
        {
            _subsitesGrid.Rows.Add(
                subsite.SiteCollectionUrl,
                subsite.SiteUrl,
                subsite.SiteTitle,
                subsite.SubsiteUrl,
                subsite.SubsiteTitle,
                subsite.WebTemplate,
                subsite.Created.ToString("g"),
                subsite.LastModified.ToString("g"),
                subsite.LanguageDisplay
            );
        }

        SetStatus($"Showing {subsites.Count} subsites across {result.SitesWithSubsites} sites | " +
                  $"Total subsites: {result.TotalSubsites} | Sites with subsites: {result.SitesWithSubsites} | " +
                  $"Sites without: {result.SitesWithoutSubsites}");
    }

    private void DisplaySiteSummary(SubsitesReportResult result)
    {
        _summaryList.Items.Clear();

        // Populate site filter dropdown with specific sites
        var currentSelection = _filterSiteCombo.SelectedIndex;
        _filterSiteCombo.Items.Clear();
        _filterSiteCombo.Items.Add("All Sites");
        _filterSiteCombo.Items.Add("Sites with Subsites Only");

        foreach (var site in result.SiteResults.OrderByDescending(s => s.SubsiteCount))
        {
            var item = new ListViewItem(site.SiteUrl)
            {
                Tag = site
            };
            item.SubItems.Add(site.SiteTitle);
            item.SubItems.Add(site.SubsiteCount.ToString());
            item.SubItems.Add(site.Success ? "Success" : "Failed");
            item.SubItems.Add(site.ErrorMessage ?? "");

            // Color code
            if (!site.Success)
            {
                item.BackColor = Color.FromArgb(255, 200, 200); // red
            }
            else if (site.HasSubsites)
            {
                item.BackColor = Color.FromArgb(255, 255, 200); // yellow - has subsites
            }
            // No color = no subsites

            _summaryList.Items.Add(item);

            // Add to filter dropdown
            if (site.HasSubsites)
            {
                _filterSiteCombo.Items.Add(site.SiteUrl);
            }
        }

        if (currentSelection >= 0 && currentSelection < _filterSiteCombo.Items.Count)
        {
            _filterSiteCombo.SelectedIndex = currentSelection;
        }
        else
        {
            _filterSiteCombo.SelectedIndex = 0;
        }
    }

    private void FilterCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplaySubsites(_currentResult);
        }
    }

    private void SearchTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplaySubsites(_currentResult);
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
        _exportSubsitesButton.Enabled = false;
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
        _subsitesGrid.Rows.Clear();
        _summaryList.Items.Clear();
        _logTextBox.Clear();
        _filterSiteCombo.SelectedIndex = 0;

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
            _currentResult = await _taskService.ExecuteSubsitesReportAsync(
                _task,
                _authService,
                progress,
                _cancellationTokenSource.Token);

            DisplayResults(_currentResult);

            if (_currentResult.Success)
            {
                SetStatus($"Task completed. {_currentResult.TotalSubsites} subsites found across {_currentResult.SitesWithSubsites} sites.");
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
            _exportSubsitesButton.Enabled = _currentResult != null;
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

    private void ExportSubsitesButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_Subsites_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportSubsitesReport(_currentResult, dialog.FileName);
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
            FileName = $"{safeName}_SiteSummary_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportSubsitesReportSummary(_currentResult, dialog.FileName);
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

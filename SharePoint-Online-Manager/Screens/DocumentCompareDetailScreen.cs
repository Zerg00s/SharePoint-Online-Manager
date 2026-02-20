using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Forms.Dialogs;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing document compare task details and results.
/// </summary>
public class DocumentCompareDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _continueButton = null!;
    private Button _exportXlsxButton = null!;
    private Button _exportCsvButton = null!;
    private Button _deleteButton = null!;
    private ListView _overviewList = null!;
    private ListView _sitesDetailList = null!;
    private ListView _issuesList = null!;
    private TextBox _logTextBox = null!;
    private TextBox _sitesDetailFilterTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;

    private TaskDefinition _task = null!;
    private DocumentCompareResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private CsvExporter _csvExporter = null!;
    private ExcelExporter _excelExporter = null!;

    private static readonly string ReportsFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SharePointOnlineManager", "Reports");

    public override string ScreenTitle => _task?.Name ?? "Document Compare Details";

    protected override void OnInitialize()
    {
        _taskService = GetRequiredService<ITaskService>();
        _authService = GetRequiredService<IAuthenticationService>();
        _connectionManager = GetRequiredService<IConnectionManager>();
        _csvExporter = GetRequiredService<CsvExporter>();
        _excelExporter = GetRequiredService<ExcelExporter>();

        // Ensure reports folder exists
        Directory.CreateDirectory(ReportsFolder);

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

        _continueButton = new Button
        {
            Text = "\u23E9 Continue",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            Enabled = false,
            BackColor = Color.FromArgb(0, 150, 80),
            ForeColor = Color.White
        };
        _continueButton.FlatAppearance.BorderColor = Color.FromArgb(0, 150, 80);
        _continueButton.FlatAppearance.BorderSize = 1;
        _continueButton.Click += ContinueButton_Click;

        _exportXlsxButton = new Button
        {
            Text = "\U0001F4CA Export XLSX",
            Size = new Size(130, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportXlsxButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportXlsxButton.FlatAppearance.BorderSize = 1;
        _exportXlsxButton.Click += ExportXlsxButton_Click;

        _exportCsvButton = new Button
        {
            Text = "\U0001F4BE Export CSV",
            Size = new Size(120, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat
        };
        _exportCsvButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportCsvButton.FlatAppearance.BorderSize = 1;
        _exportCsvButton.Click += ExportCsvButton_Click;

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

        buttonPanel.Controls.AddRange(new Control[] { _runButton, _continueButton, _exportXlsxButton, _exportCsvButton, _deleteButton });

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

        var cancelButton = new Button
        {
            Text = "Cancel",
            Location = new Point(620, 3),
            Size = new Size(80, 24),
            Name = "CancelButton"
        };
        cancelButton.Click += (s, e) => _cancellationTokenSource?.Cancel();

        progressPanel.Controls.AddRange(new Control[] { _progressBar, _progressLabel, cancelButton });

        // Tab control for results
        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        // Overview tab - high-level summary (bird's eye view)
        var overviewTab = new TabPage("Overview");
        _overviewList = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _overviewList.Columns.Add("Statistic", 300);
        _overviewList.Columns.Add("Value", 200);
        _overviewList.ContextMenuStrip = CreateListViewContextMenu(_overviewList);
        EnableCellTextSelection(_overviewList);
        overviewTab.Controls.Add(_overviewList);

        // Sites Detail tab - all sites with full stats
        var sitesDetailTab = new TabPage("Sites Detail");

        // Filter panel at top of sites detail tab
        var sitesDetailFilterPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 35
        };
        var sitesDetailFilterLabel = new Label
        {
            Text = "Filter:",
            AutoSize = true,
            Location = new Point(5, 10)
        };
        _sitesDetailFilterTextBox = new TextBox
        {
            Location = new Point(50, 7),
            Size = new Size(300, 23),
            PlaceholderText = "Type to filter by URL..."
        };
        _sitesDetailFilterTextBox.TextChanged += (s, e) =>
        {
            if (_currentResult != null)
                DisplaySitesDetail(_currentResult);
        };
        sitesDetailFilterPanel.Controls.Add(sitesDetailFilterLabel);
        sitesDetailFilterPanel.Controls.Add(_sitesDetailFilterTextBox);

        _sitesDetailList = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _sitesDetailList.Columns.Add("Source Site", 280);
        _sitesDetailList.Columns.Add("Target Site", 280);
        _sitesDetailList.Columns.Add("Src Docs", 65);
        _sitesDetailList.Columns.Add("Tgt Docs", 65);
        _sitesDetailList.Columns.Add("Found", 55);
        _sitesDetailList.Columns.Add("Size Iss", 60);
        _sitesDetailList.Columns.Add("Src Only", 60);
        _sitesDetailList.Columns.Add("Tgt Only", 60);
        _sitesDetailList.Columns.Add("Newer", 55);
        _sitesDetailList.Columns.Add("% Found", 60);
        _sitesDetailList.Columns.Add("Src Size", 80);
        _sitesDetailList.Columns.Add("Tgt Size", 80);
        _sitesDetailList.Columns.Add("Avg Src Ver", 70);
        _sitesDetailList.Columns.Add("Avg Tgt Ver", 70);
        _sitesDetailList.Columns.Add("Status", 60);
        _sitesDetailList.ContextMenuStrip = CreateSitesDetailContextMenu();
        EnableCellTextSelection(_sitesDetailList);
        EnableColumnSorting(_sitesDetailList);
        sitesDetailTab.Controls.Add(_sitesDetailList);
        sitesDetailTab.Controls.Add(sitesDetailFilterPanel);

        // Sites with Issues tab
        var issuesTab = new TabPage("Sites with Issues");
        _issuesList = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true
        };
        _issuesList.Columns.Add("Source Site", 200);
        _issuesList.Columns.Add("Target Site", 200);
        _issuesList.Columns.Add("Found", 70);
        _issuesList.Columns.Add("Size Issues", 90);
        _issuesList.Columns.Add("Source Only", 85);
        _issuesList.Columns.Add("Target Only", 85);
        _issuesList.Columns.Add("Newer", 70);
        _issuesList.Columns.Add("Error", 200);
        _issuesList.ContextMenuStrip = CreateIssuesContextMenu();
        EnableCellTextSelection(_issuesList);
        EnableColumnSorting(_issuesList);
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

        _tabControl.TabPages.AddRange(new[] { overviewTab, sitesDetailTab, issuesTab, logTab });

        Controls.Add(_tabControl);
        Controls.Add(progressPanel);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        System.Diagnostics.Debug.WriteLine("[DocCompare] *** OnNavigatedToAsync CALLED ***");

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
        System.Diagnostics.Debug.WriteLine($"[DocCompare] RefreshTaskDetailsAsync called for task: {_task.Id}");

        // Reload task from database to get latest status
        var latestTask = await _taskService.GetTaskAsync(_task.Id);
        if (latestTask != null)
        {
            _task = latestTask;
            System.Diagnostics.Debug.WriteLine($"[DocCompare] Task reloaded, Status: {_task.Status}");
        }
        else
        {
            System.Diagnostics.Debug.WriteLine($"[DocCompare] WARNING: Could not reload task from database!");
        }

        _taskNameLabel.Text = _task.Name;

        var connection = await _connectionManager.GetConnectionAsync(_task.ConnectionId);
        var connectionName = connection?.Name ?? "Unknown";

        _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Status: {_task.StatusDescription} | " +
                              $"Connection: {connectionName}";

        // Load latest result
        _currentResult = await _taskService.GetLatestDocumentCompareResultAsync(_task.Id);
        System.Diagnostics.Debug.WriteLine($"[DocCompare] Loaded result: {(_currentResult != null ? $"Found, {_currentResult.SuccessfulPairs}/{_currentResult.TotalPairsProcessed} pairs" : "NULL")}");

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportXlsxButton.Enabled = true;
            _exportCsvButton.Enabled = true;

            // Enable Continue button if task was cancelled/failed and has successful results
            var canContinue = false;
            var successfulCount = _currentResult.SiteResults.Count(s => s.Success);

            System.Diagnostics.Debug.WriteLine($"[DocCompare] Task status: {_task.Status}, Successful pairs: {successfulCount}");

            if ((_task.Status == Models.TaskStatus.Cancelled || _task.Status == Models.TaskStatus.Failed)
                && successfulCount > 0)
            {
                // Check if there are more pairs to process in the config
                if (!string.IsNullOrEmpty(_task.ConfigurationJson))
                {
                    try
                    {
                        var jsonOptions = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                        var config = System.Text.Json.JsonSerializer.Deserialize<DocumentCompareConfiguration>(_task.ConfigurationJson, jsonOptions);
                        var totalInConfig = config?.SitePairs?.Count ?? 0;
                        canContinue = successfulCount < totalInConfig;
                        System.Diagnostics.Debug.WriteLine($"[DocCompare] Config has {totalInConfig} pairs, canContinue: {canContinue}");
                    }
                    catch (Exception ex)
                    {
                        // If config parsing fails, still allow continue if we have partial results
                        System.Diagnostics.Debug.WriteLine($"[DocCompare] Config parse failed: {ex.Message}, allowing continue anyway");
                        canContinue = true;
                    }
                }
                else
                {
                    // No config but we have results - allow continue
                    System.Diagnostics.Debug.WriteLine("[DocCompare] No config but have results, allowing continue");
                    canContinue = true;
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[DocCompare] Not enabling Continue - Status: {_task.Status}, Successful: {successfulCount}");
            }

            _continueButton.Enabled = canContinue;
            System.Diagnostics.Debug.WriteLine($"[DocCompare] Continue button enabled: {canContinue}");
        }
        else
        {
            _overviewList.Items.Clear();
            _sitesDetailList.Items.Clear();
            _issuesList.Items.Clear();
            _logTextBox.Text = "No results yet. Click 'Run Task' to execute.";
            _exportXlsxButton.Enabled = false;
            _exportCsvButton.Enabled = false;
            _continueButton.Enabled = false;
        }
    }

    private void DisplayResults(DocumentCompareResult result)
    {
        DisplayOverview(result);
        DisplaySitesDetail(result);
        DisplaySitesWithIssues(result);
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);
    }

    private void DisplayOverview(DocumentCompareResult result)
    {
        _overviewList.Items.Clear();

        // Execution info section
        AddOverviewSection("EXECUTION DETAILS");
        AddOverviewItem("Executed At", result.ExecutedAt.ToString("yyyy-MM-dd HH:mm:ss"));
        AddOverviewItem("Completed At", result.CompletedAt?.ToString("yyyy-MM-dd HH:mm:ss") ?? "N/A");
        AddOverviewItem("Duration", FormatDuration(result.Duration));
        AddOverviewItem("Overall Status", result.Success ? "Success" : "Failed",
            result.Success ? Color.FromArgb(200, 255, 200) : Color.FromArgb(255, 200, 200));
        if (result.ThrottleRetryCount > 0)
        {
            AddOverviewItem("Throttle Retries", result.ThrottleRetryCount.ToString());
        }

        // Site pair stats section
        AddOverviewSection("SITE PAIR STATISTICS");
        AddOverviewItem("Total Site Pairs", result.TotalPairsProcessed.ToString());
        AddOverviewItem("Successful Pairs", result.SuccessfulPairs.ToString());
        AddOverviewItem("Failed Pairs", result.FailedPairs.ToString(),
            result.FailedPairs > 0 ? Color.FromArgb(255, 200, 200) : SystemColors.Window);

        // Document stats section
        var (found, sizeIssues, sourceOnly, targetOnly, newerAtSource) = result.GetSummary();
        var totalSourceDocs = found + sourceOnly;
        var totalTargetDocs = found + targetOnly;

        AddOverviewSection("DOCUMENT STATISTICS");
        AddOverviewItem("Total Source Documents", totalSourceDocs.ToString());
        AddOverviewItem("Total Target Documents", totalTargetDocs.ToString());
        AddOverviewItem("Documents Found", found.ToString(), Color.FromArgb(200, 255, 200));
        AddOverviewItem("Size Issues (0 bytes or <30%)", sizeIssues.ToString(),
            sizeIssues > 0 ? Color.FromArgb(255, 255, 150) : SystemColors.Window);
        AddOverviewItem("Source Only (Not Migrated)", sourceOnly.ToString(),
            sourceOnly > 0 ? Color.FromArgb(255, 200, 200) : SystemColors.Window);
        AddOverviewItem("Target Only (Extra on Target)", targetOnly.ToString(),
            targetOnly > 0 ? Color.FromArgb(255, 220, 180) : SystemColors.Window);
        AddOverviewItem("Source Newer (Stale Target)", newerAtSource.ToString(),
            newerAtSource > 0 ? Color.FromArgb(255, 255, 150) : SystemColors.Window);

        // Size stats section
        AddOverviewSection("SIZE STATISTICS");
        AddOverviewItem("Total Source Size", FormatFileSize(result.TotalSourceSizeBytes));
        AddOverviewItem("Total Target Size", FormatFileSize(result.TotalTargetSizeBytes));

        // Version stats section
        AddOverviewSection("VERSION STATISTICS");
        AddOverviewItem("Avg Source Versions/Doc", result.OverallAvgSourceVersions.ToString("F2"));
        AddOverviewItem("Avg Target Versions/Doc", result.OverallAvgTargetVersions.ToString("F2"),
            Math.Abs(result.OverallAvgSourceVersions - result.OverallAvgTargetVersions) > 0.5
                ? Color.FromArgb(255, 255, 150)
                : SystemColors.Window);

        // Migration completeness
        AddOverviewSection("MIGRATION COMPLETENESS");
        var completenessColor = result.MigrationCompletenessPercent >= 99 ? Color.FromArgb(200, 255, 200)
            : result.MigrationCompletenessPercent >= 90 ? Color.FromArgb(255, 255, 150)
            : Color.FromArgb(255, 200, 200);
        AddOverviewItem("Overall Completeness", $"{result.MigrationCompletenessPercent:F1}%", completenessColor);
    }

    private void AddOverviewSection(string title)
    {
        var item = new ListViewItem(title)
        {
            Font = new Font(_overviewList.Font, FontStyle.Bold),
            BackColor = Color.FromArgb(230, 230, 230)
        };
        item.SubItems.Add("");
        _overviewList.Items.Add(item);
    }

    private void AddOverviewItem(string stat, string value, Color? backColor = null)
    {
        var item = new ListViewItem("    " + stat);
        item.SubItems.Add(value);
        if (backColor.HasValue)
        {
            item.BackColor = backColor.Value;
        }
        _overviewList.Items.Add(item);
    }

    private void DisplaySitesDetail(DocumentCompareResult result)
    {
        _sitesDetailList.Items.Clear();

        var filterText = _sitesDetailFilterTextBox?.Text?.Trim() ?? "";
        var sites = result.SiteResults.AsEnumerable();
        if (!string.IsNullOrEmpty(filterText))
        {
            sites = sites.Where(s =>
                s.SourceSiteUrl.Contains(filterText, StringComparison.OrdinalIgnoreCase) ||
                s.TargetSiteUrl.Contains(filterText, StringComparison.OrdinalIgnoreCase));
        }
        var filteredSites = sites.ToList();

        foreach (var site in filteredSites)
        {
            var item = new ListViewItem(site.SourceSiteUrl);
            item.SubItems.Add(site.TargetSiteUrl);
            item.SubItems.Add(site.TotalSourceDocuments.ToString());
            item.SubItems.Add(site.TotalTargetDocuments.ToString());
            item.SubItems.Add(site.FoundCount.ToString());
            item.SubItems.Add(site.SizeIssueCount.ToString());
            item.SubItems.Add(site.SourceOnlyCount.ToString());
            item.SubItems.Add(site.TargetOnlyCount.ToString());
            item.SubItems.Add(site.NewerAtSourceCount.ToString());
            item.SubItems.Add($"{site.PercentFound:F1}%");
            item.SubItems.Add(FormatFileSize(site.TotalSourceSizeBytes));
            item.SubItems.Add(FormatFileSize(site.TotalTargetSizeBytes));
            item.SubItems.Add(site.AvgSourceVersions.ToString("F1"));
            item.SubItems.Add(site.AvgTargetVersions.ToString("F1"));
            item.SubItems.Add(site.Success ? "OK" : "Failed");

            // Color coding
            if (!site.Success)
            {
                item.BackColor = Color.FromArgb(255, 200, 200);
            }
            else if (site.PercentFound >= 99)
            {
                item.BackColor = Color.FromArgb(200, 255, 200);
            }
            else if (site.PercentFound >= 90)
            {
                item.BackColor = Color.FromArgb(255, 255, 150);
            }
            else if (site.SourceOnlyCount > 0)
            {
                item.BackColor = Color.FromArgb(255, 220, 180);
            }

            _sitesDetailList.Items.Add(item);
        }

        // Add totals row based on filtered sites
        if (filteredSites.Count > 0)
        {
            int totalFound = 0, totalSizeIssues = 0, totalSourceOnly = 0, totalTargetOnly = 0, totalNewerAtSource = 0;
            long totalSourceSize = 0, totalTargetSize = 0;
            foreach (var site in filteredSites)
            {
                totalFound += site.FoundCount;
                totalSizeIssues += site.SizeIssueCount;
                totalSourceOnly += site.SourceOnlyCount;
                totalTargetOnly += site.TargetOnlyCount;
                totalNewerAtSource += site.NewerAtSourceCount;
                totalSourceSize += site.TotalSourceSizeBytes;
                totalTargetSize += site.TotalTargetSizeBytes;
            }
            var grandTotalSource = totalFound + totalSourceOnly;
            var grandTotalTarget = totalFound + totalTargetOnly;
            var completeness = grandTotalSource > 0 ? (double)totalFound / grandTotalSource * 100 : 100;
            var avgSourceVer = filteredSites.SelectMany(s => s.DocumentComparisons)
                .Where(d => d.Status != DocumentCompareStatus.TargetOnly).ToList();
            var avgTargetVer = filteredSites.SelectMany(s => s.DocumentComparisons)
                .Where(d => d.Status != DocumentCompareStatus.SourceOnly).ToList();

            var totalsItem = new ListViewItem("TOTALS")
            {
                Font = new Font(_sitesDetailList.Font, FontStyle.Bold),
                BackColor = Color.FromArgb(200, 200, 200)
            };
            totalsItem.SubItems.Add("");
            totalsItem.SubItems.Add(grandTotalSource.ToString());
            totalsItem.SubItems.Add(grandTotalTarget.ToString());
            totalsItem.SubItems.Add(totalFound.ToString());
            totalsItem.SubItems.Add(totalSizeIssues.ToString());
            totalsItem.SubItems.Add(totalSourceOnly.ToString());
            totalsItem.SubItems.Add(totalTargetOnly.ToString());
            totalsItem.SubItems.Add(totalNewerAtSource.ToString());
            totalsItem.SubItems.Add($"{completeness:F1}%");
            totalsItem.SubItems.Add(FormatFileSize(totalSourceSize));
            totalsItem.SubItems.Add(FormatFileSize(totalTargetSize));
            totalsItem.SubItems.Add((avgSourceVer.Count > 0 ? avgSourceVer.Average(d => d.SourceVersionCount) : 0).ToString("F1"));
            totalsItem.SubItems.Add((avgTargetVer.Count > 0 ? avgTargetVer.Average(d => d.TargetVersionCount) : 0).ToString("F1"));
            totalsItem.SubItems.Add("");
            _sitesDetailList.Items.Add(totalsItem);
        }
    }

    private void DisplaySitesWithIssues(DocumentCompareResult result)
    {
        _issuesList.Items.Clear();

        foreach (var site in result.GetSitesWithIssues())
        {
            var item = new ListViewItem(site.SourceSiteUrl);
            item.SubItems.Add(site.TargetSiteUrl);
            item.SubItems.Add(site.FoundCount.ToString());
            item.SubItems.Add(site.SizeIssueCount.ToString());
            item.SubItems.Add(site.SourceOnlyCount.ToString());
            item.SubItems.Add(site.TargetOnlyCount.ToString());
            item.SubItems.Add(site.NewerAtSourceCount.ToString());
            item.SubItems.Add(site.ErrorMessage ?? "");

            // Color by severity
            if (!site.Success)
            {
                item.BackColor = Color.FromArgb(255, 200, 200);
            }
            else if (site.SourceOnlyCount > 0)
            {
                item.BackColor = Color.FromArgb(255, 220, 180);
            }
            else if (site.SizeIssueCount > 0 || site.NewerAtSourceCount > 0)
            {
                item.BackColor = Color.FromArgb(255, 255, 150);
            }

            _issuesList.Items.Add(item);
        }
    }

    /// <summary>
    /// Adds a single site to the Sites Detail list (for real-time updates).
    /// </summary>
    private void AddSiteToDetailList(SiteDocumentCompareResult site)
    {
        var item = new ListViewItem(site.SourceSiteUrl);
        item.SubItems.Add(site.TargetSiteUrl);
        item.SubItems.Add(site.TotalSourceDocuments.ToString());
        item.SubItems.Add(site.TotalTargetDocuments.ToString());
        item.SubItems.Add(site.FoundCount.ToString());
        item.SubItems.Add(site.SizeIssueCount.ToString());
        item.SubItems.Add(site.SourceOnlyCount.ToString());
        item.SubItems.Add(site.TargetOnlyCount.ToString());
        item.SubItems.Add(site.NewerAtSourceCount.ToString());
        item.SubItems.Add($"{site.PercentFound:F1}%");
        item.SubItems.Add(FormatFileSize(site.TotalSourceSizeBytes));
        item.SubItems.Add(FormatFileSize(site.TotalTargetSizeBytes));
        item.SubItems.Add(site.AvgSourceVersions.ToString("F1"));
        item.SubItems.Add(site.AvgTargetVersions.ToString("F1"));
        item.SubItems.Add(site.Success ? "OK" : "Failed");

        // Color coding
        if (!site.Success)
        {
            item.BackColor = Color.FromArgb(255, 200, 200);
        }
        else if (site.PercentFound >= 99)
        {
            item.BackColor = Color.FromArgb(200, 255, 200);
        }
        else if (site.PercentFound >= 90)
        {
            item.BackColor = Color.FromArgb(255, 255, 150);
        }
        else if (site.SourceOnlyCount > 0)
        {
            item.BackColor = Color.FromArgb(255, 220, 180);
        }

        _sitesDetailList.Items.Add(item);
    }

    /// <summary>
    /// Adds a single site to the Sites with Issues list (for real-time updates).
    /// </summary>
    private void AddSiteToIssuesList(SiteDocumentCompareResult site)
    {
        var item = new ListViewItem(site.SourceSiteUrl);
        item.SubItems.Add(site.TargetSiteUrl);
        item.SubItems.Add(site.FoundCount.ToString());
        item.SubItems.Add(site.SizeIssueCount.ToString());
        item.SubItems.Add(site.SourceOnlyCount.ToString());
        item.SubItems.Add(site.TargetOnlyCount.ToString());
        item.SubItems.Add(site.NewerAtSourceCount.ToString());
        item.SubItems.Add(site.ErrorMessage ?? "");

        // Color by severity
        if (!site.Success)
        {
            item.BackColor = Color.FromArgb(255, 200, 200);
        }
        else if (site.SourceOnlyCount > 0)
        {
            item.BackColor = Color.FromArgb(255, 220, 180);
        }
        else if (site.SizeIssueCount > 0 || site.NewerAtSourceCount > 0)
        {
            item.BackColor = Color.FromArgb(255, 255, 150);
        }

        _issuesList.Items.Add(item);
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

    private async void ContinueButton_Click(object? sender, EventArgs e)
    {
        await ExecuteTaskAsync(continueFromPrevious: true);
    }

    private async Task ExecuteTaskAsync(bool continueFromPrevious = false)
    {
        // Setup for execution
        _runButton.Text = "Cancel";
        _runButton.BackColor = Color.DarkRed;
        _deleteButton.Enabled = false;
        _continueButton.Enabled = false;
        _exportXlsxButton.Enabled = false;
        _exportCsvButton.Enabled = false;

        var progressPanel = Controls.Find("ProgressPanel", false).FirstOrDefault();
        if (progressPanel != null)
        {
            progressPanel.Visible = true;
        }
        _progressBar.Value = 0;
        _progressLabel.Text = continueFromPrevious ? "Continuing from previous run..." : "Starting...";

        _cancellationTokenSource = new CancellationTokenSource();

        // Only clear lists if starting fresh (not continuing)
        if (!continueFromPrevious)
        {
            _overviewList.Items.Clear();
            _sitesDetailList.Items.Clear();
            _issuesList.Items.Clear();
            _logTextBox.Clear();
            _currentResult = null;
        }

        // Initialize a temporary result to track progress in real-time
        // This allows export to work even while task is running
        var progressResult = _currentResult ?? new DocumentCompareResult { TaskId = _task.Id };

        var progress = new Progress<TaskProgress>(p =>
        {
            _progressBar.Value = p.PercentComplete;
            _progressLabel.Text = p.Message;

            // Real-time update: add completed site to the lists and result
            if (p.CompletedSiteResult != null)
            {
                AddSiteToDetailList(p.CompletedSiteResult);
                if (p.CompletedSiteResult.HasIssues)
                {
                    AddSiteToIssuesList(p.CompletedSiteResult);
                }

                // Also add to the progress result so export works during execution
                if (!progressResult.SiteResults.Any(s =>
                    s.SourceSiteUrl.Equals(p.CompletedSiteResult.SourceSiteUrl, StringComparison.OrdinalIgnoreCase)))
                {
                    progressResult.SiteResults.Add(p.CompletedSiteResult);
                }
                _currentResult = progressResult;
            }
        });

        try
        {
            _currentResult = await _taskService.ExecuteDocumentCompareAsync(
                _task,
                _authService,
                _connectionManager,
                progress,
                _cancellationTokenSource.Token,
                continueFromPrevious,
                reauthCallback: async (tenantName, tenantDomain) =>
                {
                    AuthCookies? freshCookies = null;
                    Invoke(() =>
                    {
                        var siteUrl = $"https://{tenantDomain}";
                        using var loginForm = new LoginForm(siteUrl, $"Re-authenticate: {tenantDomain}");
                        var dialogResult = loginForm.ShowDialog(FindForm());
                        if (dialogResult == DialogResult.OK && loginForm.CapturedCookies != null)
                        {
                            _authService.StoreCookies(loginForm.CapturedCookies);
                            freshCookies = loginForm.CapturedCookies;
                        }
                    });
                    return freshCookies;
                });

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
            _runButton.Text = "\u25B6 Run Task";
            _runButton.BackColor = Color.FromArgb(0, 120, 212);
            _deleteButton.Enabled = true;
            _exportXlsxButton.Enabled = _currentResult != null;
            _exportCsvButton.Enabled = _currentResult != null;
            if (progressPanel != null)
            {
                progressPanel.Visible = false;
            }
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;

            // Refresh task details to show updated status and Continue button state
            var updatedTask = await _taskService.GetTaskAsync(_task.Id);
            if (updatedTask != null)
            {
                _task = updatedTask;
                _taskInfoLabel.Text = $"Type: {_task.TypeDescription} | Status: {_task.StatusDescription}";

                // Enable Continue button if task was cancelled/failed and has partial results
                var canContinue = false;
                if (_task.Status == Models.TaskStatus.Cancelled || _task.Status == Models.TaskStatus.Failed)
                {
                    if (!string.IsNullOrEmpty(_task.ConfigurationJson) && _currentResult != null)
                    {
                        try
                        {
                            var jsonOptions = new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                            var config = System.Text.Json.JsonSerializer.Deserialize<DocumentCompareConfiguration>(_task.ConfigurationJson, jsonOptions);
                            if (config != null)
                            {
                                var completedCount = _currentResult.SiteResults.Count(s => s.Success);
                                canContinue = completedCount > 0 && completedCount < config.SitePairs.Count;
                            }
                        }
                        catch { /* ignore deserialization errors */ }
                    }
                }
                _continueButton.Enabled = canContinue;
            }
        }
    }

    private void ExportXlsxButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var defaultFileName = $"{safeName}_Overview_{timestamp}.xlsx";

        using var dialog = new SaveFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx",
            FileName = defaultFileName,
            InitialDirectory = ReportsFolder
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _excelExporter.ExportDocumentCompareReport(_currentResult, dialog.FileName);
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

    private void ExportCsvButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var defaultFileName = $"{safeName}_Details_{timestamp}.csv";

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = defaultFileName,
            InitialDirectory = ReportsFolder
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportDocumentCompareReport(_currentResult, dialog.FileName);
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

    private static string FormatDuration(TimeSpan duration)
    {
        if (duration.TotalHours >= 1)
        {
            return $"{(int)duration.TotalHours}h {duration.Minutes}m {duration.Seconds}s";
        }
        if (duration.TotalMinutes >= 1)
        {
            return $"{duration.Minutes}m {duration.Seconds}s";
        }
        return $"{duration.Seconds}s";
    }

    private static string FormatFileSize(long bytes)
    {
        string[] sizes = ["B", "KB", "MB", "GB", "TB"];
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len /= 1024;
        }
        return $"{len:F2} {sizes[order]}";
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

    /// <summary>
    /// Creates a context menu for copying data from a ListView.
    /// </summary>
    private static ContextMenuStrip CreateListViewContextMenu(ListView listView)
    {
        var contextMenu = new ContextMenuStrip();

        var copyCell = new ToolStripMenuItem("Copy Cell");
        copyCell.Click += (s, e) =>
        {
            if (listView.SelectedItems.Count > 0)
            {
                var point = listView.PointToClient(Cursor.Position);
                var hitTest = listView.HitTest(point);
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
            if (listView.SelectedItems.Count > 0)
            {
                var item = listView.SelectedItems[0];
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
            foreach (ListViewItem item in listView.Items)
            {
                // Try to find URL-like columns (typically first two columns for source/target)
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

        contextMenu.Items.Add(copyCell);
        contextMenu.Items.Add(copyRow);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(copyAllUrls);

        return contextMenu;
    }

    /// <summary>
    /// Creates a context menu for the Sites Detail list with export option.
    /// </summary>
    private ContextMenuStrip CreateSitesDetailContextMenu()
    {
        var contextMenu = new ContextMenuStrip();

        var copyCell = new ToolStripMenuItem("Copy Cell");
        copyCell.Click += (s, e) =>
        {
            if (_sitesDetailList.SelectedItems.Count > 0)
            {
                var point = _sitesDetailList.PointToClient(Cursor.Position);
                var hitTest = _sitesDetailList.HitTest(point);
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
            if (_sitesDetailList.SelectedItems.Count > 0)
            {
                var item = _sitesDetailList.SelectedItems[0];
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
            foreach (ListViewItem item in _sitesDetailList.Items)
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

        var exportSiteCsv = new ToolStripMenuItem("Export Site to CSV");
        exportSiteCsv.Click += ExportSiteToCsv_Click;

        var previewIssues = new ToolStripMenuItem("Preview Issues");
        previewIssues.Click += (s, e) => PreviewSiteIssues(_sitesDetailList);

        contextMenu.Items.Add(copyCell);
        contextMenu.Items.Add(copyRow);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(copyAllUrls);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(exportSiteCsv);
        contextMenu.Items.Add(previewIssues);

        // Enable/disable items based on selection
        contextMenu.Opening += (s, e) =>
        {
            var hasSelection = _sitesDetailList.SelectedItems.Count > 0;
            var isNotTotalsRow = hasSelection &&
                _sitesDetailList.SelectedItems[0].Text != "TOTALS";
            exportSiteCsv.Enabled = hasSelection && isNotTotalsRow && _currentResult != null;

            var siteResult = hasSelection && isNotTotalsRow && _currentResult != null
                ? FindSiteResult(_sitesDetailList.SelectedItems[0].Text)
                : null;
            previewIssues.Enabled = siteResult != null && siteResult.HasIssues;
        };

        return contextMenu;
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

        var previewIssues = new ToolStripMenuItem("Preview Issues");
        previewIssues.Click += (s, e) => PreviewSiteIssues(_issuesList);

        var exportIssuesCsv = new ToolStripMenuItem("Export Issues to CSV");
        exportIssuesCsv.Click += ExportIssuesToCsv_Click;

        contextMenu.Items.Add(copyCell);
        contextMenu.Items.Add(copyRow);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(copyAllUrls);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add(previewIssues);
        contextMenu.Items.Add(exportIssuesCsv);

        contextMenu.Opening += (s, e) =>
        {
            var hasSelection = _issuesList.SelectedItems.Count > 0;
            previewIssues.Enabled = hasSelection && _currentResult != null;
            exportIssuesCsv.Enabled = _currentResult != null;
        };

        return contextMenu;
    }

    private SiteDocumentCompareResult? FindSiteResult(string sourceUrl)
    {
        return _currentResult?.SiteResults.FirstOrDefault(
            s => s.SourceSiteUrl.Equals(sourceUrl, StringComparison.OrdinalIgnoreCase));
    }

    private void PreviewSiteIssues(ListView listView)
    {
        if (listView.SelectedItems.Count == 0 || _currentResult == null)
            return;

        var sourceUrl = listView.SelectedItems[0].Text;
        if (sourceUrl == "TOTALS")
            return;

        var siteResult = FindSiteResult(sourceUrl);
        if (siteResult == null || !siteResult.HasIssues)
            return;

        using var dialog = new IssuesPreviewDialog(siteResult, _csvExporter);
        dialog.ShowDialog(FindForm());
    }

    private void ExportIssuesToCsv_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var answer = MessageBox.Show(
            "Include documents without issues from these sites?\n\n" +
            "Yes = all documents from sites with issues\n" +
            "No = only issue rows (Source Only, Size Issue, Newer at Source)",
            "Export Issues",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (answer == DialogResult.Cancel)
            return;

        var issuesSites = _currentResult.GetSitesWithIssues().ToList();
        IEnumerable<DocumentCompareItem> items;

        if (answer == DialogResult.Yes)
        {
            items = issuesSites.SelectMany(s => s.DocumentComparisons);
        }
        else
        {
            items = issuesSites.SelectMany(s => s.DocumentComparisons)
                .Where(d => d.Status == DocumentCompareStatus.SourceOnly ||
                            d.Status == DocumentCompareStatus.SizeIssue ||
                            d.IsNewerAtSource);
        }

        var safeName = SanitizeFileName(_task.Name);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var defaultFileName = $"{safeName}_Issues_{timestamp}.csv";

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = defaultFileName,
            InitialDirectory = ReportsFolder
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportDocumentCompareItems(items, dialog.FileName);
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

    private void ExportSiteToCsv_Click(object? sender, EventArgs e)
    {
        if (_sitesDetailList.SelectedItems.Count == 0 || _currentResult == null)
            return;

        var selectedItem = _sitesDetailList.SelectedItems[0];
        var sourceUrl = selectedItem.Text;

        // Skip if TOTALS row is selected
        if (sourceUrl == "TOTALS")
            return;

        // Find the matching site result
        var siteResult = _currentResult.SiteResults.FirstOrDefault(
            s => s.SourceSiteUrl.Equals(sourceUrl, StringComparison.OrdinalIgnoreCase));

        if (siteResult == null)
        {
            MessageBox.Show("Could not find site data for export.", "Export Error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Generate a safe filename from the site URL
        var siteName = new Uri(sourceUrl).AbsolutePath.Trim('/').Replace("/", "_");
        if (string.IsNullOrEmpty(siteName))
        {
            siteName = new Uri(sourceUrl).Host.Replace(".", "_");
        }
        var safeSiteName = SanitizeFileName(siteName);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var defaultFileName = $"DocCompare_{safeSiteName}_{timestamp}.csv";

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = defaultFileName,
            InitialDirectory = AppContext.BaseDirectory
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportSiteDocumentCompareReport(siteResult, dialog.FileName);
                SetStatus($"Exported {siteResult.TotalDocuments} documents to {dialog.FileName}");
                OfferToOpenFile(dialog.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

/// <summary>
/// Comparer for sorting ListView columns. Handles numeric, percentage, and file size values.
/// </summary>
internal class ListViewColumnComparer : System.Collections.IComparer
{
    private readonly int _column;
    private readonly SortOrder _order;

    public ListViewColumnComparer(int column, SortOrder order)
    {
        _column = column;
        _order = order;
    }

    public int Compare(object? x, object? y)
    {
        if (x is not ListViewItem itemX || y is not ListViewItem itemY)
            return 0;

        // Keep TOTALS row always at the bottom
        if (itemX.Text == "TOTALS") return 1;
        if (itemY.Text == "TOTALS") return -1;

        var textX = _column < itemX.SubItems.Count ? itemX.SubItems[_column].Text : "";
        var textY = _column < itemY.SubItems.Count ? itemY.SubItems[_column].Text : "";

        int result;

        // Try numeric comparison (handles plain numbers)
        if (double.TryParse(textX, out var numX) && double.TryParse(textY, out var numY))
        {
            result = numX.CompareTo(numY);
        }
        // Try percentage comparison (e.g., "95.5%")
        else if (textX.EndsWith('%') && textY.EndsWith('%') &&
                 double.TryParse(textX.TrimEnd('%'), out var pctX) &&
                 double.TryParse(textY.TrimEnd('%'), out var pctY))
        {
            result = pctX.CompareTo(pctY);
        }
        // Try file size comparison (e.g., "1.25 GB", "500.00 MB")
        else if (TryParseFileSize(textX, out var sizeX) && TryParseFileSize(textY, out var sizeY))
        {
            result = sizeX.CompareTo(sizeY);
        }
        else
        {
            result = string.Compare(textX, textY, StringComparison.OrdinalIgnoreCase);
        }

        return _order == SortOrder.Descending ? -result : result;
    }

    private static bool TryParseFileSize(string text, out double bytes)
    {
        bytes = 0;
        var parts = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2 || !double.TryParse(parts[0], out var value))
            return false;

        var multiplier = parts[1].ToUpperInvariant() switch
        {
            "B" => 1.0,
            "KB" => 1024.0,
            "MB" => 1024.0 * 1024,
            "GB" => 1024.0 * 1024 * 1024,
            "TB" => 1024.0 * 1024 * 1024 * 1024,
            _ => -1.0
        };

        if (multiplier < 0) return false;
        bytes = value * multiplier;
        return true;
    }
}

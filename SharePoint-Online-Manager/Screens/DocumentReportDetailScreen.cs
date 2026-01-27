using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for viewing document report task details and results.
/// </summary>
public class DocumentReportDetailScreen : BaseScreen
{
    private Label _taskNameLabel = null!;
    private Label _taskInfoLabel = null!;
    private Button _runButton = null!;
    private Button _exportButton = null!;
    private Button _exportSummaryButton = null!;
    private Button _deleteButton = null!;
    private DataGridView _documentsGrid = null!;
    private ListView _summaryList = null!;
    private TextBox _logTextBox = null!;
    private TabControl _tabControl = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private ComboBox _filterCombo = null!;
    private TextBox _searchTextBox = null!;

    private TaskDefinition _task = null!;
    private DocumentReportResult? _currentResult;
    private CancellationTokenSource? _cancellationTokenSource;
    private ITaskService _taskService = null!;
    private IAuthenticationService _authService = null!;
    private IConnectionManager _connectionManager = null!;
    private CsvExporter _csvExporter = null!;

    public override string ScreenTitle => _task?.Name ?? "Document Report Details";

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
            Text = "\U0001F4BE Export All",
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

        // Tab control for documents, summary, and log
        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        // Documents tab
        var documentsTab = new TabPage("Documents");

        var documentsHeaderPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 35,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        var filterLabel = new Label
        {
            Text = "Filter by site:",
            AutoSize = true,
            Padding = new Padding(0, 5, 5, 0)
        };

        _filterCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 300
        };
        _filterCombo.Items.Add("All Sites");
        _filterCombo.SelectedIndex = 0;
        _filterCombo.SelectedIndexChanged += FilterCombo_SelectedIndexChanged;

        var searchLabel = new Label
        {
            Text = "Search:",
            AutoSize = true,
            Padding = new Padding(20, 5, 5, 0)
        };

        _searchTextBox = new TextBox
        {
            Width = 200,
            PlaceholderText = "Search file name..."
        };
        _searchTextBox.TextChanged += SearchTextBox_TextChanged;

        documentsHeaderPanel.Controls.AddRange(new Control[] { filterLabel, _filterCombo, searchLabel, _searchTextBox });

        _documentsGrid = new DataGridView
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
            if (_documentsGrid.CurrentCell != null)
            {
                Clipboard.SetText(_documentsGrid.CurrentCell.Value?.ToString() ?? "");
            }
        });
        copyMenuItem.ShortcutKeys = Keys.Control | Keys.C;
        var copyRowMenuItem = new ToolStripMenuItem("Copy Row", null, (s, e) =>
        {
            if (_documentsGrid.CurrentRow != null)
            {
                var values = new List<string>();
                foreach (DataGridViewCell cell in _documentsGrid.CurrentRow.Cells)
                {
                    values.Add(cell.Value?.ToString() ?? "");
                }
                Clipboard.SetText(string.Join("\t", values));
            }
        });
        gridContextMenu.Items.Add(copyMenuItem);
        gridContextMenu.Items.Add(copyRowMenuItem);
        _documentsGrid.ContextMenuStrip = gridContextMenu;

        _documentsGrid.Columns.Add("FileName", "File Name");
        _documentsGrid.Columns.Add("Extension", "Ext");
        _documentsGrid.Columns.Add("Size", "Size");
        _documentsGrid.Columns.Add("Created", "Created");
        _documentsGrid.Columns.Add("CreatedBy", "Created By");
        _documentsGrid.Columns.Add("Modified", "Modified");
        _documentsGrid.Columns.Add("ModifiedBy", "Modified By");
        _documentsGrid.Columns.Add("Library", "Library");
        _documentsGrid.Columns.Add("Site", "Site");
        _documentsGrid.Columns.Add("Versions", "Ver");

        // Adjust column widths
        _documentsGrid.Columns["FileName"].FillWeight = 120;
        _documentsGrid.Columns["Extension"].FillWeight = 40;
        _documentsGrid.Columns["Size"].FillWeight = 60;
        _documentsGrid.Columns["Created"].FillWeight = 80;
        _documentsGrid.Columns["CreatedBy"].FillWeight = 80;
        _documentsGrid.Columns["Modified"].FillWeight = 80;
        _documentsGrid.Columns["ModifiedBy"].FillWeight = 80;
        _documentsGrid.Columns["Library"].FillWeight = 80;
        _documentsGrid.Columns["Site"].FillWeight = 100;
        _documentsGrid.Columns["Versions"].FillWeight = 40;

        documentsTab.Controls.Add(_documentsGrid);
        documentsTab.Controls.Add(documentsHeaderPanel);

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
        _summaryList.Columns.Add("Site Title", 150);
        _summaryList.Columns.Add("Libraries", 80);
        _summaryList.Columns.Add("Documents", 100);
        _summaryList.Columns.Add("Total Size", 100);
        _summaryList.Columns.Add("Status", 80);
        _summaryList.Columns.Add("Error", 200);

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

        _tabControl.TabPages.AddRange(new[] { documentsTab, summaryTab, logTab });

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
                              $"Connection: {connectionName} | Sites: {_task.TotalSites}";

        // Load latest result
        _currentResult = await _taskService.GetLatestDocumentReportResultAsync(_task.Id);

        if (_currentResult != null)
        {
            DisplayResults(_currentResult);
            _exportButton.Enabled = true;
            _exportSummaryButton.Enabled = true;
        }
        else
        {
            _documentsGrid.Rows.Clear();
            _summaryList.Items.Clear();
            _logTextBox.Text = "No results yet. Click 'Run Task' to execute.";
            _exportButton.Enabled = false;
            _exportSummaryButton.Enabled = false;
        }
    }

    private void DisplayResults(DocumentReportResult result)
    {
        PopulateSiteFilter(result);
        DisplayDocuments(result);
        DisplaySiteSummary(result);
        _logTextBox.Text = string.Join(Environment.NewLine, result.ExecutionLog);
    }

    private void PopulateSiteFilter(DocumentReportResult result)
    {
        _filterCombo.Items.Clear();
        _filterCombo.Items.Add("All Sites");

        foreach (var site in result.SiteResults.Where(s => s.Success))
        {
            _filterCombo.Items.Add(site.SiteUrl);
        }

        _filterCombo.SelectedIndex = 0;
    }

    private void DisplayDocuments(DocumentReportResult result)
    {
        _documentsGrid.Rows.Clear();

        var documents = result.GetAllDocuments().ToList();

        // Apply site filter
        var selectedSite = _filterCombo.SelectedIndex > 0 ? _filterCombo.SelectedItem?.ToString() : null;
        if (!string.IsNullOrEmpty(selectedSite))
        {
            documents = documents.Where(d => d.SiteCollectionUrl.Equals(selectedSite, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        // Apply search filter
        var searchText = _searchTextBox.Text.Trim();
        if (!string.IsNullOrEmpty(searchText))
        {
            documents = documents.Where(d =>
                d.FileName.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                d.Extension.Contains(searchText, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        foreach (var doc in documents)
        {
            _documentsGrid.Rows.Add(
                doc.FileName,
                doc.Extension,
                doc.SizeFormatted,
                doc.CreatedDate.ToString("yyyy-MM-dd HH:mm"),
                doc.CreatedBy,
                doc.ModifiedDate.ToString("yyyy-MM-dd HH:mm"),
                doc.ModifiedBy,
                doc.LibraryTitle,
                doc.SiteTitle,
                doc.VersionCount
            );
        }

        var (totalDocs, totalSize, _) = result.GetSummary();
        SetStatus($"Showing {documents.Count} of {totalDocs} documents ({FormatSize(totalSize)} total)");
    }

    private void DisplaySiteSummary(DocumentReportResult result)
    {
        _summaryList.Items.Clear();

        foreach (var site in result.SiteResults)
        {
            var item = new ListViewItem(site.SiteUrl);
            item.SubItems.Add(site.SiteTitle);
            item.SubItems.Add(site.LibrariesProcessed.ToString());
            item.SubItems.Add(site.TotalDocuments.ToString());
            item.SubItems.Add(FormatSize(site.TotalSizeBytes));
            item.SubItems.Add(site.Success ? "Success" : "Failed");
            item.SubItems.Add(site.ErrorMessage ?? "");

            if (!site.Success)
            {
                item.BackColor = Color.FromArgb(255, 200, 200);
            }
            else if (site.TotalDocuments == 0)
            {
                item.BackColor = Color.FromArgb(255, 255, 200);
            }

            _summaryList.Items.Add(item);
        }
    }

    private void FilterCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayDocuments(_currentResult);
        }
    }

    private void SearchTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (_currentResult != null)
        {
            DisplayDocuments(_currentResult);
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
        _exportButton.Enabled = false;
        _exportSummaryButton.Enabled = false;

        var progressPanel = Controls.Find("ProgressPanel", false).FirstOrDefault();
        if (progressPanel != null)
        {
            progressPanel.Visible = true;
        }
        _progressBar.Value = 0;
        _progressLabel.Text = "Starting...";

        _cancellationTokenSource = new CancellationTokenSource();
        _documentsGrid.Rows.Clear();
        _summaryList.Items.Clear();
        _logTextBox.Clear();
        _filterCombo.Items.Clear();
        _filterCombo.Items.Add("All Sites");
        _filterCombo.SelectedIndex = 0;

        var progress = new Progress<TaskProgress>(p =>
        {
            _progressBar.Value = p.PercentComplete;
            _progressLabel.Text = p.Message;
        });

        try
        {
            _currentResult = await _taskService.ExecuteDocumentReportAsync(
                _task,
                _authService,
                progress,
                _cancellationTokenSource.Token);

            DisplayResults(_currentResult);

            if (_currentResult.Success)
            {
                var (totalDocs, totalSize, totalLibs) = _currentResult.GetSummary();
                SetStatus($"Task completed. {totalDocs} documents, {FormatSize(totalSize)}, {totalLibs} libraries.");
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
            _runButton.Text = "Run Task";
            _deleteButton.Enabled = true;
            _exportButton.Enabled = _currentResult != null;
            _exportSummaryButton.Enabled = _currentResult != null;
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

    private void ExportButton_Click(object? sender, EventArgs e)
    {
        if (_currentResult == null)
            return;

        var safeName = SanitizeFileName(_task.Name);
        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"{safeName}_Documents_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportDocumentReport(_currentResult, dialog.FileName);
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
                _csvExporter.ExportDocumentReportSummary(_currentResult, dialog.FileName);
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

    private static string FormatSize(long bytes)
    {
        string[] sizes = ["B", "KB", "MB", "GB", "TB"];
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }
}

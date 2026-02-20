using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Detail screen for a tenant pair showing site mappings and task trigger buttons.
/// </summary>
public class TenantPairDetailScreen : BaseScreen
{
    private ListView _sitePairsListView = null!;
    private TextBox _filterTextBox = null!;
    private Button _importCsvButton = null!;
    private Button _clearAllButton = null!;
    private Button _compareDocsButton = null!;
    private Button _compareListsButton = null!;
    private Button _siteAccessButton = null!;
    private Button _navSettingsButton = null!;
    private Label _headerLabel = null!;
    private Label _connectionInfoLabel = null!;

    private ITenantPairService _tenantPairService = null!;
    private IConnectionManager _connectionManager = null!;
    private TenantPair _pair = null!;
    private int _sortColumn = -1;
    private SortOrder _sortOrder = SortOrder.None;

    public override string ScreenTitle => "Tenant Pair Details";

    protected override void OnInitialize()
    {
        _tenantPairService = GetRequiredService<ITenantPairService>();
        _connectionManager = GetRequiredService<IConnectionManager>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();

        // Header label
        _headerLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 28,
            Font = new Font(Font.FontFamily, 11, FontStyle.Bold),
            Padding = new Padding(5, 5, 0, 0)
        };

        _connectionInfoLabel = new Label
        {
            Dock = DockStyle.Top,
            Height = 22,
            ForeColor = SystemColors.GrayText,
            Padding = new Padding(5, 0, 0, 0)
        };

        // Button panel
        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 42,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(3, 5, 0, 5)
        };

        _importCsvButton = new Button
        {
            Text = "Import CSV",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 8, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _importCsvButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _importCsvButton.Click += ImportCsvButton_Click;

        _clearAllButton = new Button
        {
            Text = "Clear All",
            Size = new Size(80, 28),
            Margin = new Padding(0, 0, 20, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat,
            ForeColor = Color.DarkRed
        };
        _clearAllButton.FlatAppearance.BorderColor = Color.DarkRed;
        _clearAllButton.Click += ClearAllButton_Click;

        // Separator
        var separator = new Label
        {
            Text = "|",
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
            Margin = new Padding(0, 6, 10, 0)
        };

        _compareDocsButton = new Button
        {
            Text = "Compare Documents",
            Size = new Size(145, 28),
            Margin = new Padding(0, 0, 8, 0),
            FlatStyle = FlatStyle.Flat
        };
        _compareDocsButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _compareDocsButton.Click += CompareDocsButton_Click;

        _compareListsButton = new Button
        {
            Text = "Compare Lists",
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 8, 0),
            FlatStyle = FlatStyle.Flat
        };
        _compareListsButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _compareListsButton.Click += CompareListsButton_Click;

        _siteAccessButton = new Button
        {
            Text = "Site Access Check",
            Size = new Size(130, 28),
            Margin = new Padding(0, 0, 8, 0),
            FlatStyle = FlatStyle.Flat
        };
        _siteAccessButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _siteAccessButton.Click += SiteAccessButton_Click;

        _navSettingsButton = new Button
        {
            Text = "Nav Settings Sync",
            Size = new Size(130, 28),
            Margin = new Padding(0, 0, 0, 0),
            FlatStyle = FlatStyle.Flat
        };
        _navSettingsButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _navSettingsButton.Click += NavSettingsButton_Click;

        buttonPanel.Controls.AddRange(new Control[]
        {
            _importCsvButton, _clearAllButton, separator,
            _compareDocsButton, _compareListsButton, _siteAccessButton, _navSettingsButton
        });

        // Filter panel
        var filterPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 32,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(3, 3, 0, 3)
        };

        var filterLabel = new Label
        {
            Text = "Filter:",
            AutoSize = true,
            Margin = new Padding(0, 5, 5, 0)
        };

        _filterTextBox = new TextBox
        {
            Size = new Size(300, 23),
            PlaceholderText = "Type to filter by URL..."
        };
        _filterTextBox.TextChanged += FilterTextBox_TextChanged;

        filterPanel.Controls.AddRange(new Control[] { filterLabel, _filterTextBox });

        // Site pairs ListView
        _sitePairsListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = true
        };
        _sitePairsListView.Columns.Add("Source URL", 350);
        _sitePairsListView.Columns.Add("Target URL", 350);
        _sitePairsListView.ColumnClick += SitePairsListView_ColumnClick;

        // Context menu
        var contextMenu = new ContextMenuStrip();
        var copyCell = new ToolStripMenuItem("Copy Cell");
        copyCell.Click += CopyCell_Click;
        var copyRow = new ToolStripMenuItem("Copy Row");
        copyRow.Click += CopyRow_Click;
        var copyAll = new ToolStripMenuItem("Copy All URLs");
        copyAll.Click += CopyAllUrls_Click;
        var deleteSelected = new ToolStripMenuItem("Delete Selected");
        deleteSelected.Click += DeleteSelected_Click;
        contextMenu.Items.AddRange(new ToolStripItem[] { copyCell, copyRow, copyAll, new ToolStripSeparator(), deleteSelected });
        _sitePairsListView.ContextMenuStrip = contextMenu;

        // Add controls in reverse dock order (Fill last)
        Controls.Add(_sitePairsListView);
        Controls.Add(filterPanel);
        Controls.Add(buttonPanel);
        Controls.Add(_connectionInfoLabel);
        Controls.Add(_headerLabel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        if (parameter is TenantPair pair)
        {
            _pair = pair;
        }

        if (_pair == null)
            return;

        // Load connection names for display
        var connections = await _connectionManager.GetAllConnectionsAsync();
        var sourceConn = connections.FirstOrDefault(c => c.Id == _pair.SourceConnectionId);
        var targetConn = connections.FirstOrDefault(c => c.Id == _pair.TargetConnectionId);

        _headerLabel.Text = _pair.Name ?? "Tenant Pair";
        _connectionInfoLabel.Text = $"{sourceConn?.Name ?? "(deleted)"} \u2192 {targetConn?.Name ?? "(deleted)"}";

        DisplaySitePairs();
    }

    private void DisplaySitePairs()
    {
        _sitePairsListView.BeginUpdate();
        try
        {
            _sitePairsListView.Items.Clear();
            var filterText = _filterTextBox.Text.Trim();

            var pairs = _pair.SitePairs.AsEnumerable();
            if (!string.IsNullOrEmpty(filterText))
            {
                pairs = pairs.Where(p =>
                    p.SourceUrl.Contains(filterText, StringComparison.OrdinalIgnoreCase) ||
                    p.TargetUrl.Contains(filterText, StringComparison.OrdinalIgnoreCase));
            }

            foreach (var sp in pairs)
            {
                var item = new ListViewItem(sp.SourceUrl) { Tag = sp };
                item.SubItems.Add(sp.TargetUrl);
                _sitePairsListView.Items.Add(item);
            }

            _clearAllButton.Enabled = _pair.SitePairs.Count > 0;
            SetStatus($"{_sitePairsListView.Items.Count} site pair(s)" +
                      (_pair.SitePairs.Count != _sitePairsListView.Items.Count
                          ? $" (filtered from {_pair.SitePairs.Count})"
                          : ""));
        }
        finally
        {
            _sitePairsListView.EndUpdate();
        }
    }

    private void FilterTextBox_TextChanged(object? sender, EventArgs e)
    {
        DisplaySitePairs();
    }

    private void SitePairsListView_ColumnClick(object? sender, ColumnClickEventArgs e)
    {
        if (e.Column == _sortColumn)
        {
            _sortOrder = _sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
        }
        else
        {
            _sortColumn = e.Column;
            _sortOrder = SortOrder.Ascending;
        }

        _sitePairsListView.ListViewItemSorter = new ListViewItemComparer(_sortColumn, _sortOrder);
        _sitePairsListView.Sort();
    }

    private void ImportCsvButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
            Title = "Select Site Pairs CSV File"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        try
        {
            var lines = File.ReadAllLines(dialog.FileName);
            var importedPairs = new List<SiteComparePair>();
            var errors = new List<string>();
            var lineNumber = 0;

            foreach (var line in lines)
            {
                lineNumber++;

                if (string.IsNullOrWhiteSpace(line))
                    continue;

                var parts = line.Split(',');
                if (parts.Length < 2)
                {
                    if (lineNumber == 1 && (line.Contains("Source", StringComparison.OrdinalIgnoreCase) ||
                                            line.Contains("URL", StringComparison.OrdinalIgnoreCase)))
                        continue;

                    errors.Add($"Line {lineNumber}: Not enough columns");
                    continue;
                }

                var sourceUrl = parts[0].Trim().Trim('"');
                var targetUrl = parts[1].Trim().Trim('"');

                if (!Uri.TryCreate(sourceUrl, UriKind.Absolute, out _))
                {
                    if (lineNumber == 1) continue;
                    errors.Add($"Line {lineNumber}: Invalid source URL '{sourceUrl}'");
                    continue;
                }

                if (!Uri.TryCreate(targetUrl, UriKind.Absolute, out _))
                {
                    if (lineNumber == 1) continue;
                    errors.Add($"Line {lineNumber}: Invalid target URL '{targetUrl}'");
                    continue;
                }

                importedPairs.Add(new SiteComparePair
                {
                    SourceUrl = sourceUrl,
                    TargetUrl = targetUrl
                });
            }

            if (errors.Count > 0 && importedPairs.Count == 0)
            {
                MessageBox.Show(
                    $"Failed to import CSV file:\n\n{string.Join("\n", errors.Take(10))}",
                    "Import Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            _pair.SitePairs.AddRange(importedPairs);
            SavePairAsync();
            DisplaySitePairs();

            var message = $"Imported {importedPairs.Count} site pair(s).";
            if (errors.Count > 0)
            {
                message += $"\n\n{errors.Count} line(s) had errors and were skipped.";
                MessageBox.Show(
                    $"{message}\n\nErrors:\n{string.Join("\n", errors.Take(5))}",
                    "Import Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            SetStatus(message);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to read CSV file: {ex.Message}",
                "Import Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void ClearAllButton_Click(object? sender, EventArgs e)
    {
        var result = MessageBox.Show(
            "Are you sure you want to clear all site pairs?",
            "Clear All",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            _pair.SitePairs.Clear();
            SavePairAsync();
            DisplaySitePairs();
            SetStatus("All site pairs cleared");
        }
    }

    private void DeleteSelected_Click(object? sender, EventArgs e)
    {
        if (_sitePairsListView.SelectedItems.Count == 0)
            return;

        var toRemove = _sitePairsListView.SelectedItems
            .Cast<ListViewItem>()
            .Select(item => (SiteComparePair)item.Tag)
            .ToList();

        foreach (var pair in toRemove)
        {
            _pair.SitePairs.Remove(pair);
        }

        SavePairAsync();
        DisplaySitePairs();
        SetStatus($"Removed {toRemove.Count} site pair(s)");
    }

    private void CopyCell_Click(object? sender, EventArgs e)
    {
        if (_sitePairsListView.SelectedItems.Count == 0)
            return;

        var item = _sitePairsListView.SelectedItems[0];
        // Determine which column was clicked based on cursor position
        var point = _sitePairsListView.PointToClient(Cursor.Position);
        var hitTest = _sitePairsListView.HitTest(point);
        var text = hitTest.SubItem?.Text ?? item.Text;
        if (!string.IsNullOrEmpty(text))
            Clipboard.SetText(text);
    }

    private void CopyRow_Click(object? sender, EventArgs e)
    {
        if (_sitePairsListView.SelectedItems.Count == 0)
            return;

        var lines = _sitePairsListView.SelectedItems
            .Cast<ListViewItem>()
            .Select(item => $"{item.Text}\t{item.SubItems[1].Text}");
        Clipboard.SetText(string.Join(Environment.NewLine, lines));
    }

    private void CopyAllUrls_Click(object? sender, EventArgs e)
    {
        if (_sitePairsListView.Items.Count == 0)
            return;

        var lines = _sitePairsListView.Items
            .Cast<ListViewItem>()
            .Select(item => $"{item.Text},{item.SubItems[1].Text}");
        Clipboard.SetText("Source URL,Target URL" + Environment.NewLine + string.Join(Environment.NewLine, lines));
    }

    private async void CompareDocsButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.NavigateToAsync<DocumentCompareConfigScreen>(
            new TenantPairTaskContext { TenantPair = _pair });
    }

    private async void CompareListsButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.NavigateToAsync<ListCompareConfigScreen>(
            new TenantPairTaskContext { TenantPair = _pair });
    }

    private async void SiteAccessButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.NavigateToAsync<SiteAccessConfigScreen>(
            new TenantPairTaskContext { TenantPair = _pair });
    }

    private async void NavSettingsButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.NavigateToAsync<NavigationSettingsConfigScreen>(
            new TenantPairTaskContext { TenantPair = _pair });
    }

    private async void SavePairAsync()
    {
        try
        {
            await _tenantPairService.SavePairAsync(_pair);
        }
        catch (Exception ex)
        {
            SetStatus($"Failed to save: {ex.Message}");
        }
    }
}

/// <summary>
/// Comparer for sorting ListView columns.
/// </summary>
internal class ListViewItemComparer : System.Collections.IComparer
{
    private readonly int _column;
    private readonly SortOrder _order;

    public ListViewItemComparer(int column, SortOrder order)
    {
        _column = column;
        _order = order;
    }

    public int Compare(object? x, object? y)
    {
        var itemX = (ListViewItem)x!;
        var itemY = (ListViewItem)y!;
        var textX = _column == 0 ? itemX.Text : itemX.SubItems[_column].Text;
        var textY = _column == 0 ? itemY.Text : itemY.SubItems[_column].Text;
        var result = string.Compare(textX, textY, StringComparison.OrdinalIgnoreCase);
        return _order == SortOrder.Descending ? -result : result;
    }
}

using SharePointOnlineManager.Models;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Forms.Dialogs;

/// <summary>
/// Modal dialog that shows issue documents for a single site in a DataGridView.
/// </summary>
public class IssuesPreviewDialog : Form
{
    private readonly SiteDocumentCompareResult _siteResult;
    private readonly CsvExporter _csvExporter;
    private readonly List<DocumentCompareItem> _issueItems;
    private DataGridView _grid = null!;

    public IssuesPreviewDialog(SiteDocumentCompareResult siteResult, CsvExporter csvExporter)
    {
        _siteResult = siteResult;
        _csvExporter = csvExporter;

        _issueItems = siteResult.DocumentComparisons
            .Where(d => d.Status == DocumentCompareStatus.SourceOnly ||
                        d.Status == DocumentCompareStatus.SizeIssue ||
                        d.IsNewerAtSource)
            .ToList();

        InitializeComponent();
        LoadData();
    }

    private void InitializeComponent()
    {
        Text = $"Issues Preview \u2014 {_siteResult.SourceSiteUrl}";
        Size = new Size(1100, 600);
        MinimumSize = new Size(800, 400);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = false;
        ShowInTaskbar = false;

        // Header panel
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 40,
            Padding = new Padding(8, 8, 8, 4)
        };

        var sourceOnlyCount = _issueItems.Count(d => d.Status == DocumentCompareStatus.SourceOnly);
        var sizeIssueCount = _issueItems.Count(d => d.Status == DocumentCompareStatus.SizeIssue);
        var newerCount = _issueItems.Count(d => d.IsNewerAtSource);

        var headerLabel = new Label
        {
            Text = $"{_siteResult.SourceSiteUrl}  —  {sourceOnlyCount} Source Only, {sizeIssueCount} Size Issues, {newerCount} Newer at Source",
            AutoSize = true,
            Location = new Point(8, 10),
            Font = new Font(Font.FontFamily, 9F)
        };
        headerPanel.Controls.Add(headerLabel);

        // DataGridView
        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            BackgroundColor = Color.White,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText,
            AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(245, 245, 245) },
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        };

        AddColumn("Library", 120);
        AddColumn("File Name", 180);
        AddColumn("Status", 90);
        AddColumn("Newer?", 55);
        AddColumn("Source Size", 90);
        AddColumn("Target Size", 90);
        AddColumn("Size Diff %", 70);
        AddColumn("Source Path", 200);
        AddColumn("Target Path", 200);
        AddColumn("Src Modified", 120);
        AddColumn("Tgt Modified", 120);

        // Context menu for grid
        var gridContextMenu = new ContextMenuStrip();
        var copyCellItem = new ToolStripMenuItem("Copy Cell");
        copyCellItem.Click += (s, e) =>
        {
            if (_grid.CurrentCell != null)
            {
                var text = _grid.CurrentCell.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(text))
                    Clipboard.SetText(text);
            }
        };
        gridContextMenu.Items.Add(copyCellItem);
        _grid.ContextMenuStrip = gridContextMenu;

        // Button panel
        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(10, 8, 10, 8)
        };

        var closeButton = new Button
        {
            Text = "Close",
            Size = new Size(90, 30),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White
        };
        closeButton.FlatAppearance.BorderColor = Color.FromArgb(90, 98, 104);
        closeButton.Click += (s, e) => Close();

        var exportButton = new Button
        {
            Text = "Export to CSV",
            Size = new Size(110, 30),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        exportButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        exportButton.Click += ExportButton_Click;

        buttonPanel.Controls.Add(closeButton);
        buttonPanel.Controls.Add(exportButton);

        // Add controls (order matters for docking)
        Controls.Add(_grid);
        Controls.Add(buttonPanel);
        Controls.Add(headerPanel);

        CancelButton = closeButton;
    }

    private void AddColumn(string headerText, int width)
    {
        var col = new DataGridViewTextBoxColumn
        {
            HeaderText = headerText,
            Width = width,
            SortMode = DataGridViewColumnSortMode.Automatic
        };
        _grid.Columns.Add(col);
    }

    private void LoadData()
    {
        _grid.Rows.Clear();

        foreach (var doc in _issueItems)
        {
            var rowIndex = _grid.Rows.Add(
                doc.LibraryName,
                doc.FileName,
                doc.StatusDescription,
                doc.IsNewerAtSource ? "Yes" : "",
                FormatFileSize(doc.SourceSizeBytes),
                FormatFileSize(doc.TargetSizeBytes),
                $"{doc.SizeDifferencePercent:F1}%",
                doc.SourceAbsolutePath,
                doc.TargetAbsolutePath,
                doc.SourceModified?.ToString("yyyy-MM-dd HH:mm") ?? "",
                doc.TargetModified?.ToString("yyyy-MM-dd HH:mm") ?? ""
            );

            var row = _grid.Rows[rowIndex];

            // Color-code rows by issue type
            if (doc.Status == DocumentCompareStatus.SourceOnly)
            {
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 200, 200);
            }
            else if (doc.Status == DocumentCompareStatus.SizeIssue)
            {
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 220, 180);
            }
            else if (doc.IsNewerAtSource)
            {
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 150);
            }
        }
    }

    private void ExportButton_Click(object? sender, EventArgs e)
    {
        var siteName = "site";
        try
        {
            var uri = new Uri(_siteResult.SourceSiteUrl);
            siteName = uri.AbsolutePath.Trim('/').Replace("/", "_");
            if (string.IsNullOrEmpty(siteName))
                siteName = uri.Host.Replace(".", "_");
        }
        catch { /* use default */ }

        var safeName = SanitizeFileName(siteName);
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"Issues_{safeName}_{timestamp}.csv"
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            try
            {
                _csvExporter.ExportDocumentCompareItems(_issueItems, dialog.FileName);
                MessageBox.Show($"Exported {_issueItems.Count} issue(s) to:\n{dialog.FileName}",
                    "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
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

    private static string SanitizeFileName(string fileName)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        foreach (var c in invalidChars)
            fileName = fileName.Replace(c, '_');
        return fileName;
    }
}

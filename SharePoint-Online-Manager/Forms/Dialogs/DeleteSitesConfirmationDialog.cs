using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Forms.Dialogs;

/// <summary>
/// Confirmation dialog for permanently deleting site collections.
/// </summary>
public class DeleteSitesConfirmationDialog : Form
{
    private readonly List<SiteCollection> _sites;
    private DataGridView _sitesGrid = null!;
    private Label _warningLabel = null!;
    private Button _cancelButton = null!;
    private Button _deleteButton = null!;

    public DeleteSitesConfirmationDialog(List<SiteCollection> sites)
    {
        _sites = sites;
        InitializeComponent();
        LoadSites();
    }

    private void InitializeComponent()
    {
        Text = "Confirm Permanent Deletion";
        Size = new Size(700, 500);
        MinimumSize = new Size(500, 400);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;

        // Warning panel at the top
        var warningPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 80,
            BackColor = Color.FromArgb(255, 243, 205),
            Padding = new Padding(15)
        };

        var warningIcon = new Label
        {
            Text = "\u26A0", // Warning sign
            Font = new Font("Segoe UI", 24, FontStyle.Bold),
            ForeColor = Color.FromArgb(133, 100, 4),
            AutoSize = true,
            Location = new Point(15, 20)
        };

        _warningLabel = new Label
        {
            Text = $"WARNING: You are about to permanently delete {_sites.Count} site collection(s).\n" +
                   "This action will send the sites to the SharePoint Recycle Bin.\n" +
                   "Please review the sites below before proceeding.",
            Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
            ForeColor = Color.FromArgb(133, 100, 4),
            AutoSize = true,
            Location = new Point(70, 15),
            MaximumSize = new Size(580, 0)
        };

        warningPanel.Controls.Add(warningIcon);
        warningPanel.Controls.Add(_warningLabel);

        // Sites grid
        _sitesGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            MultiSelect = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = Color.White
        };

        _sitesGrid.Columns.Add("Title", "Site Title");
        _sitesGrid.Columns.Add("Url", "Site URL");
        _sitesGrid.Columns.Add("SiteType", "Type");

        _sitesGrid.Columns["Title"].FillWeight = 30;
        _sitesGrid.Columns["Url"].FillWeight = 50;
        _sitesGrid.Columns["SiteType"].FillWeight = 20;

        // Button panel
        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(10)
        };

        _deleteButton = new Button
        {
            Text = "Yes, permanently delete",
            Size = new Size(180, 35),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9.5F, FontStyle.Bold)
        };
        _deleteButton.FlatAppearance.BorderColor = Color.FromArgb(200, 35, 51);
        _deleteButton.Click += DeleteButton_Click;

        _cancelButton = new Button
        {
            Text = "Cancel",
            Size = new Size(100, 35),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(108, 117, 125),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9.5F)
        };
        _cancelButton.FlatAppearance.BorderColor = Color.FromArgb(90, 98, 104);
        _cancelButton.Click += CancelButton_Click;

        buttonPanel.Controls.Add(_deleteButton);
        buttonPanel.Controls.Add(_cancelButton);

        // Add controls in order
        Controls.Add(_sitesGrid);
        Controls.Add(buttonPanel);
        Controls.Add(warningPanel);

        // Set cancel button as the cancel result
        CancelButton = _cancelButton;
    }

    private void LoadSites()
    {
        foreach (var site in _sites)
        {
            _sitesGrid.Rows.Add(
                site.Title,
                site.Url,
                site.SiteTypeDescription
            );
        }
    }

    private void CancelButton_Click(object? sender, EventArgs e)
    {
        DialogResult = DialogResult.Cancel;
        Close();
    }

    private void DeleteButton_Click(object? sender, EventArgs e)
    {
        DialogResult = DialogResult.Yes;
        Close();
    }
}

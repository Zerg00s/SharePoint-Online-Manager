using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Forms.Dialogs;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for displaying and selecting site collections from a SharePoint tenant.
/// </summary>
public class SiteCollectionsScreen : BaseScreen
{
    private TabControl _tabControl = null!;
    private DataGridView _sitesGrid = null!;
    private TreeView _hubsTreeView = null!;
    private ComboBox _filterComboBox = null!;
    private ComboBox _standaloneFilter = null!;
    private ComboBox _groupFilter = null!;
    private ComboBox _channelFilter = null!;
    private ComboBox _sharingFilter = null!;
    private ComboBox _stateFilter = null!;
    private TextBox _searchTextBox = null!;
    private Button _refreshButton = null!;
    private Button _importUrlsButton = null!;
    private Button _exportButton = null!;
    private Button _exploreButton = null!;
    private Button _createTaskButton = null!;
    private Button _deleteButton = null!;
    private Button _selectAllButton = null!;
    private Button _selectNoneButton = null!;
    private Label _countLabel = null!;

    private Connection _connection = null!;
    private List<SiteCollection> _allSites = [];
    private List<SiteCollection> _filteredSites = [];
    private IAuthenticationService _authService = null!;
    private int _sortColumnIndex = -1;
    private SortOrder _sortOrder = SortOrder.None;

    private string _screenTitle = "Site Collections";
    public override string ScreenTitle => _screenTitle;

    protected override void OnInitialize()
    {
        _authService = GetRequiredService<IAuthenticationService>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();

        // Top toolbar panel
        var toolbarPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 45,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        var filterLabel = new Label
        {
            Text = "Filter:",
            AutoSize = true,
            Margin = new Padding(0, 8, 5, 0)
        };

        _filterComboBox = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 150,
            Margin = new Padding(0, 0, 15, 0)
        };
        _filterComboBox.Items.AddRange(["All Sites", "SharePoint Sites Only", "OneDrive Only"]);
        _filterComboBox.SelectedIndex = 0;
        _filterComboBox.SelectedIndexChanged += FilterComboBox_SelectedIndexChanged;

        var searchLabel = new Label
        {
            Text = "Search:",
            AutoSize = true,
            Margin = new Padding(0, 8, 5, 0)
        };

        _searchTextBox = new TextBox
        {
            Width = 200,
            Margin = new Padding(0, 0, 15, 0)
        };
        _searchTextBox.TextChanged += SearchTextBox_TextChanged;

        _refreshButton = new Button
        {
            Text = "\U0001F504 Refresh", // Refresh emoji
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat
        };
        _refreshButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _refreshButton.FlatAppearance.BorderSize = 1;
        _refreshButton.Click += RefreshButton_Click;

        _importUrlsButton = new Button
        {
            Text = "\U0001F4C2 Import URLs", // File folder emoji
            Size = new Size(120, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat
        };
        _importUrlsButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _importUrlsButton.FlatAppearance.BorderSize = 1;
        _importUrlsButton.Click += ImportUrlsButton_Click;

        _exportButton = new Button
        {
            Text = "\U0001F4BE Export to CSV", // Floppy disk emoji
            Size = new Size(130, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            Enabled = false
        };
        _exportButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exportButton.FlatAppearance.BorderSize = 1;
        _exportButton.Click += ExportButton_Click;

        _exploreButton = new Button
        {
            Text = "\U0001F50D Explore Site", // Magnifying glass emoji
            Size = new Size(120, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat,
            Enabled = false
        };
        _exploreButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _exploreButton.FlatAppearance.BorderSize = 1;
        _exploreButton.Click += ExploreButton_Click;

        toolbarPanel.Controls.AddRange(new Control[]
        {
            filterLabel, _filterComboBox,
            searchLabel, _searchTextBox,
            _refreshButton, _importUrlsButton, _exportButton, _exploreButton
        });

        // Second filter row
        var filterPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 35,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 2, 0, 2)
        };

        var standaloneLabel = new Label { Text = "Standalone:", AutoSize = true, Margin = new Padding(0, 6, 3, 0) };
        _standaloneFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60, Margin = new Padding(0, 0, 10, 0) };
        _standaloneFilter.Items.AddRange(["All", "Yes", "No"]);
        _standaloneFilter.SelectedIndex = 0;
        _standaloneFilter.SelectedIndexChanged += Filter_Changed;

        var groupLabel = new Label { Text = "Group:", AutoSize = true, Margin = new Padding(0, 6, 3, 0) };
        _groupFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60, Margin = new Padding(0, 0, 10, 0) };
        _groupFilter.Items.AddRange(["All", "Yes", "No"]);
        _groupFilter.SelectedIndex = 0;
        _groupFilter.SelectedIndexChanged += Filter_Changed;

        var channelLabel = new Label { Text = "Channel:", AutoSize = true, Margin = new Padding(0, 6, 3, 0) };
        _channelFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 60, Margin = new Padding(0, 0, 10, 0) };
        _channelFilter.Items.AddRange(["All", "Yes", "No"]);
        _channelFilter.SelectedIndex = 0;
        _channelFilter.SelectedIndexChanged += Filter_Changed;

        var sharingLabel = new Label { Text = "Sharing:", AutoSize = true, Margin = new Padding(0, 6, 3, 0) };
        _sharingFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70, Margin = new Padding(0, 0, 10, 0) };
        _sharingFilter.Items.AddRange(["All", "On", "Off"]);
        _sharingFilter.SelectedIndex = 0;
        _sharingFilter.SelectedIndexChanged += Filter_Changed;

        var stateLabel = new Label { Text = "State:", AutoSize = true, Margin = new Padding(0, 6, 3, 0) };
        _stateFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 80, Margin = new Padding(0, 0, 10, 0) };
        _stateFilter.Items.AddRange(["All", "Active", "Locked"]);
        _stateFilter.SelectedIndex = 0;
        _stateFilter.SelectedIndexChanged += Filter_Changed;

        filterPanel.Controls.AddRange(new Control[]
        {
            standaloneLabel, _standaloneFilter,
            groupLabel, _groupFilter,
            channelLabel, _channelFilter,
            sharingLabel, _sharingFilter,
            stateLabel, _stateFilter
        });

        // Bottom action panel
        var actionPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 50,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 10, 0, 5)
        };

        _selectAllButton = new Button
        {
            Text = "\u2611 Select All", // Checked box
            Size = new Size(110, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat
        };
        _selectAllButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _selectAllButton.FlatAppearance.BorderSize = 1;
        _selectAllButton.Click += SelectAllButton_Click;

        _selectNoneButton = new Button
        {
            Text = "\u2610 Select None", // Empty box
            Size = new Size(115, 28),
            Margin = new Padding(0, 0, 20, 0),
            FlatStyle = FlatStyle.Flat
        };
        _selectNoneButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _selectNoneButton.FlatAppearance.BorderSize = 1;
        _selectNoneButton.Click += SelectNoneButton_Click;

        _createTaskButton = new Button
        {
            Text = "\u2795 Create Task from Selection", // Plus sign
            Size = new Size(200, 28),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        _createTaskButton.FlatAppearance.BorderColor = Color.FromArgb(0, 120, 212);
        _createTaskButton.FlatAppearance.BorderSize = 1;
        _createTaskButton.Click += CreateTaskButton_Click;

        _deleteButton = new Button
        {
            Text = "\U0001F5D1 Delete Permanently", // Wastebasket emoji
            Size = new Size(160, 28),
            Margin = new Padding(0, 0, 20, 0),
            Enabled = false,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White
        };
        _deleteButton.FlatAppearance.BorderColor = Color.FromArgb(200, 35, 51);
        _deleteButton.FlatAppearance.BorderSize = 1;
        _deleteButton.Click += DeleteButton_Click;

        _countLabel = new Label
        {
            AutoSize = true,
            Margin = new Padding(20, 8, 0, 0)
        };

        actionPanel.Controls.AddRange(new Control[]
        {
            _selectAllButton, _selectNoneButton, _createTaskButton, _deleteButton, _countLabel
        });

        // Sites DataGridView
        _sitesGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            MultiSelect = true,
            ReadOnly = false,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        // Add columns
        var selectColumn = new DataGridViewCheckBoxColumn
        {
            Name = "Select",
            HeaderText = "",
            Width = 30,
            AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        };
        _sitesGrid.Columns.Add(selectColumn);
        _sitesGrid.Columns.Add("Title", "Title");
        _sitesGrid.Columns.Add("Url", "URL");
        _sitesGrid.Columns.Add("SiteType", "Type");
        _sitesGrid.Columns.Add("Template", "Template");
        _sitesGrid.Columns.Add("StorageUsed", "Storage Used");
        _sitesGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "StorageUsedBytes",
            HeaderText = "Storage (Bytes)",
            Visible = false
        });
        _sitesGrid.Columns.Add("Owner", "Owner");
        _sitesGrid.Columns.Add("Standalone", "Standalone");
        _sitesGrid.Columns.Add("Group", "Group");
        _sitesGrid.Columns.Add("Channel", "Channel");
        _sitesGrid.Columns.Add("Hub", "Hub");
        _sitesGrid.Columns.Add("Files", "Files");
        _sitesGrid.Columns.Add("PageViews", "Views");
        _sitesGrid.Columns.Add("LastActivity", "Last Activity");
        _sitesGrid.Columns.Add("ExternalSharing", "Sharing");
        _sitesGrid.Columns.Add("State", "State");
        _sitesGrid.Columns.Add("Language", "Language");

        _sitesGrid.CellValueChanged += SitesGrid_CellValueChanged;
        _sitesGrid.CurrentCellDirtyStateChanged += SitesGrid_CurrentCellDirtyStateChanged;
        _sitesGrid.ColumnHeaderMouseClick += SitesGrid_ColumnHeaderMouseClick;
        _sitesGrid.SelectionChanged += SitesGrid_SelectionChanged;
        _sitesGrid.CellDoubleClick += SitesGrid_CellDoubleClick;

        // Hubs TreeView
        _hubsTreeView = new TreeView
        {
            Dock = DockStyle.Fill,
            Margin = Padding.Empty,
            ShowLines = true,
            ShowPlusMinus = true,
            ShowRootLines = true,
            HideSelection = false,
            Scrollable = true,
            ShowNodeToolTips = true,
            ImageList = CreateHubImageList()
        };
        _hubsTreeView.NodeMouseDoubleClick += HubsTreeView_NodeMouseDoubleClick;

        // Tab Control
        _tabControl = new TabControl
        {
            Dock = DockStyle.Fill
        };

        var sitesTab = new TabPage("Sites") { UseVisualStyleBackColor = true, Padding = Padding.Empty };
        sitesTab.Controls.Add(_sitesGrid);

        var hubsTab = new TabPage("Hub Sites") { UseVisualStyleBackColor = true, Padding = Padding.Empty };
        hubsTab.Controls.Add(_hubsTreeView);

        _tabControl.TabPages.Add(sitesTab);
        _tabControl.TabPages.Add(hubsTab);

        Controls.Add(_tabControl);
        Controls.Add(actionPanel);
        Controls.Add(filterPanel);
        Controls.Add(toolbarPanel);

        ResumeLayout(true);
    }

    private static ImageList CreateHubImageList()
    {
        var imageList = new ImageList { ImageSize = new Size(16, 16), ColorDepth = ColorDepth.Depth32Bit };

        // 0: Hub site (home icon - blue)
        imageList.Images.Add("hub", CreateIcon(Color.FromArgb(0, 120, 212), IconShape.Home));
        // 1: Group/Teams site (people icon - purple)
        imageList.Images.Add("group", CreateIcon(Color.FromArgb(128, 0, 128), IconShape.People));
        // 2: Channel site (chat icon - teal)
        imageList.Images.Add("channel", CreateIcon(Color.FromArgb(0, 128, 128), IconShape.Chat));
        // 3: Regular site (document icon - gray)
        imageList.Images.Add("site", CreateIcon(Color.FromArgb(100, 100, 100), IconShape.Document));
        // 4: Folder (for standalone section)
        imageList.Images.Add("folder", CreateIcon(Color.FromArgb(200, 150, 50), IconShape.Folder));

        return imageList;
    }

    private enum IconShape { Home, People, Chat, Document, Folder }

    private static Bitmap CreateIcon(Color color, IconShape shape)
    {
        var bmp = new Bitmap(16, 16);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        g.Clear(Color.Transparent);

        using var brush = new SolidBrush(color);
        using var pen = new Pen(color, 1.5f);

        switch (shape)
        {
            case IconShape.Home:
                // House shape
                var housePoints = new Point[] {
                    new(8, 1), new(15, 7), new(12, 7), new(12, 14),
                    new(4, 14), new(4, 7), new(1, 7)
                };
                g.FillPolygon(brush, housePoints);
                break;

            case IconShape.People:
                // Two circles for people
                g.FillEllipse(brush, 3, 2, 5, 5);
                g.FillEllipse(brush, 8, 2, 5, 5);
                g.FillEllipse(brush, 1, 8, 6, 6);
                g.FillEllipse(brush, 9, 8, 6, 6);
                break;

            case IconShape.Chat:
                // Chat bubble
                g.FillEllipse(brush, 1, 1, 14, 10);
                var tailPoints = new Point[] { new(3, 10), new(1, 14), new(6, 10) };
                g.FillPolygon(brush, tailPoints);
                break;

            case IconShape.Document:
                // Document with folded corner
                g.FillRectangle(brush, 3, 1, 10, 14);
                using (var whiteBrush = new SolidBrush(Color.White))
                {
                    var foldPoints = new Point[] { new(9, 1), new(13, 5), new(9, 5) };
                    g.FillPolygon(whiteBrush, foldPoints);
                    // Lines for text
                    g.FillRectangle(whiteBrush, 5, 7, 6, 1);
                    g.FillRectangle(whiteBrush, 5, 10, 6, 1);
                }
                break;

            case IconShape.Folder:
                // Folder shape
                g.FillRectangle(brush, 1, 4, 14, 10);
                g.FillRectangle(brush, 1, 2, 6, 3);
                break;
        }

        return bmp;
    }

    private void HubsTreeView_NodeMouseDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
    {
        if (e.Node?.Tag is SiteCollection site)
        {
            // Navigate to site explorer or select in grid
            _tabControl.SelectedIndex = 0; // Switch to Sites tab

            // Find and select the site in the grid
            foreach (DataGridViewRow row in _sitesGrid.Rows)
            {
                if (row.Tag is SiteCollection rowSite && rowSite.Url == site.Url)
                {
                    row.Selected = true;
                    _sitesGrid.FirstDisplayedScrollingRowIndex = row.Index;
                    break;
                }
            }
        }
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        if (parameter is Connection connection)
        {
            _connection = connection;
            _screenTitle = $"Site Collections - {_connection.Name}";
            UpdateTitle();
        }
        else if (_connection == null)
        {
            MessageBox.Show("No connection specified.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            await NavigationService!.GoBackAsync();
            return;
        }

        await LoadSitesAsync();
    }

    private async Task LoadSitesAsync()
    {
        ShowLoading("Loading site collections...");
        _sitesGrid.Rows.Clear();
        _allSites.Clear();

        try
        {
            var cookies = _authService.GetStoredCookies(_connection.CookieDomain);
            if (cookies == null || !cookies.IsValid)
            {
                var authenticated = await ReauthenticateAsync();
                if (!authenticated)
                {
                    HideLoading();
                    return;
                }
                cookies = _authService.GetStoredCookies(_connection.CookieDomain);
            }

            if (cookies == null)
            {
                throw new InvalidOperationException("No valid credentials available.");
            }

            using var adminService = new AdminService(cookies, _connection.TenantName);

            var progress = new Progress<string>(message => SetStatus(message));
            _allSites = await adminService.GetAllSiteCollectionsAsync(progress);

            ApplyFilter();
            SetStatus($"Loaded {_allSites.Count} site collections");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Import URLs"))
        {
            // Search API not available - show helpful message
            MessageBox.Show(
                ex.Message,
                "Site Enumeration Not Available",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            SetStatus("Use 'Import URLs' to add sites manually");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load site collections: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Failed to load sites - use 'Import URLs' to add sites manually");
        }
        finally
        {
            HideLoading();
        }
    }

    private async Task<bool> ReauthenticateAsync()
    {
        var result = MessageBox.Show(
            "Authentication expired or missing. Would you like to sign in?",
            "Authentication Required",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result != DialogResult.Yes)
        {
            return false;
        }

        using var loginForm = new LoginForm(_connection.PrimaryUrl);
        if (loginForm.ShowDialog(FindForm()) == DialogResult.OK && loginForm.CapturedCookies != null)
        {
            _authService.StoreCookies(loginForm.CapturedCookies);
            return true;
        }

        return false;
    }

    private void ApplyFilter()
    {
        var searchText = _searchTextBox.Text.Trim().ToLowerInvariant();
        var filterIndex = _filterComboBox.SelectedIndex;
        var standaloneFilter = _standaloneFilter.SelectedIndex;
        var groupFilter = _groupFilter.SelectedIndex;
        var channelFilter = _channelFilter.SelectedIndex;
        var sharingFilter = _sharingFilter.SelectedIndex;
        var stateFilter = _stateFilter.SelectedIndex;

        _filteredSites = _allSites.Where(site =>
        {
            // Apply type filter
            if (filterIndex == 1 && site.SiteType == SiteType.OneDrive)
                return false;
            if (filterIndex == 2 && site.SiteType != SiteType.OneDrive)
                return false;

            // Apply standalone filter
            var isStandalone = !site.IsGroupConnected && site.ChannelType == 0;
            if (standaloneFilter == 1 && !isStandalone) return false;
            if (standaloneFilter == 2 && isStandalone) return false;

            // Apply group filter
            if (groupFilter == 1 && !site.IsGroupConnected) return false;
            if (groupFilter == 2 && site.IsGroupConnected) return false;

            // Apply channel filter
            if (channelFilter == 1 && site.ChannelType == 0) return false;
            if (channelFilter == 2 && site.ChannelType != 0) return false;

            // Apply sharing filter
            if (sharingFilter == 1 && !site.ExternalSharing.Equals("On", StringComparison.OrdinalIgnoreCase)) return false;
            if (sharingFilter == 2 && !site.ExternalSharing.Equals("Off", StringComparison.OrdinalIgnoreCase)) return false;

            // Apply state filter (0 and 1 are Active, 2 is Locked)
            if (stateFilter == 1 && site.State > 1) return false;
            if (stateFilter == 2 && site.State <= 1) return false;

            // Apply search filter
            if (!string.IsNullOrEmpty(searchText))
            {
                return site.Title.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                       site.Url.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                       site.Owner.Contains(searchText, StringComparison.OrdinalIgnoreCase);
            }

            return true;
        }).ToList();

        RefreshGrid();
    }

    private void RefreshGrid()
    {
        if (_sitesGrid.Columns.Count == 0)
            return;

        _sitesGrid.Rows.Clear();

        foreach (var site in _filteredSites)
        {
            var isStandalone = !site.IsGroupConnected && site.ChannelType == 0;
            var rowIndex = _sitesGrid.Rows.Add(
                false,
                site.Title,
                site.Url,
                site.SiteTypeDescription,
                site.Template,
                site.StorageUsedFormatted,
                site.StorageUsedBytes, // Hidden column for sorting
                site.Owner,
                isStandalone ? "Yes" : "No",
                site.IsGroupConnected ? "Yes" : "No",
                site.ChannelTypeDisplay,
                site.HubDisplay,
                site.FileCount,
                site.PageViews,
                site.LastActivityDisplay,
                site.ExternalSharing,
                site.StateDisplay,
                site.LanguageDisplay
            );
            _sitesGrid.Rows[rowIndex].Tag = site;
        }

        _exportButton.Enabled = _filteredSites.Count > 0;
        UpdateCountLabel();
        RefreshHubsTree();
    }

    private void UpdateCountLabel()
    {
        var selectedCount = GetSelectedSites().Count;
        _countLabel.Text = $"{selectedCount} of {_filteredSites.Count} sites selected (Total: {_allSites.Count})";
        _createTaskButton.Enabled = selectedCount > 0;
        _deleteButton.Enabled = selectedCount > 0;
    }

    private void RefreshHubsTree()
    {
        _hubsTreeView.Nodes.Clear();

        // Find all hub sites (sites where SiteId == HubSiteId)
        var hubSites = _allSites
            .Where(s => s.SiteId != Guid.Empty && s.SiteId == s.HubSiteId)
            .OrderBy(s => s.Title)
            .ToList();

        // Find standalone sites (not associated with any hub)
        var standaloneSites = _allSites
            .Where(s => s.HubSiteId == Guid.Empty)
            .OrderBy(s => s.Title)
            .ToList();

        // Create hub nodes with their associated sites
        foreach (var hub in hubSites)
        {
            // Find sites associated with this hub (excluding the hub itself)
            var associatedSites = _allSites
                .Where(s => s.HubSiteId == hub.SiteId && s.SiteId != hub.SiteId)
                .OrderBy(s => s.Title)
                .ToList();

            var hubNode = new TreeNode($"[{associatedSites.Count}] {hub.Title} ({GetSitePath(hub.Url)})")
            {
                Tag = hub,
                ImageKey = "hub",
                SelectedImageKey = "hub",
                ToolTipText = hub.Url
            };

            foreach (var site in associatedSites)
            {
                var imageKey = site.ChannelType > 0 ? "channel" :
                              site.IsGroupConnected ? "group" : "site";
                var siteNode = new TreeNode($"{site.Title} ({GetSitePath(site.Url)})")
                {
                    Tag = site,
                    ImageKey = imageKey,
                    SelectedImageKey = imageKey,
                    ToolTipText = site.Url
                };
                hubNode.Nodes.Add(siteNode);
            }

            _hubsTreeView.Nodes.Add(hubNode);
        }

        // Add standalone sites section if there are any
        if (standaloneSites.Count > 0)
        {
            var standaloneNode = new TreeNode($"[{standaloneSites.Count}] Standalone Sites")
            {
                ImageKey = "folder",
                SelectedImageKey = "folder"
            };

            foreach (var site in standaloneSites.Take(100)) // Limit to first 100 for performance
            {
                var icon = site.ChannelType > 0 ? "\U0001F4AC" :
                          site.IsGroupConnected ? "\U0001F465" :
                          "\U0001F4C4";
                var siteNode = new TreeNode($"{icon} {site.Title} ({GetSitePath(site.Url)})")
                {
                    Tag = site,
                    ToolTipText = site.Url
                };
                standaloneNode.Nodes.Add(siteNode);
            }

            if (standaloneSites.Count > 100)
            {
                standaloneNode.Nodes.Add(new TreeNode($"... and {standaloneSites.Count - 100} more"));
            }

            _hubsTreeView.Nodes.Add(standaloneNode);
        }

        // Expand hub nodes by default
        foreach (TreeNode node in _hubsTreeView.Nodes)
        {
            if (node.Tag is SiteCollection) // It's a hub, not standalone section
            {
                node.Expand();
            }
        }
    }

    private List<SiteCollection> GetSelectedSites()
    {
        var selected = new List<SiteCollection>();
        foreach (DataGridViewRow row in _sitesGrid.Rows)
        {
            if (row.Cells["Select"].Value is true && row.Tag is SiteCollection site)
            {
                selected.Add(site);
            }
        }
        return selected;
    }

    private void FilterComboBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        ApplyFilter();
    }

    private void Filter_Changed(object? sender, EventArgs e)
    {
        ApplyFilter();
    }

    private void SearchTextBox_TextChanged(object? sender, EventArgs e)
    {
        ApplyFilter();
    }

    private async void RefreshButton_Click(object? sender, EventArgs e)
    {
        await LoadSitesAsync();
    }

    private void ImportUrlsButton_Click(object? sender, EventArgs e)
    {
        using var dialog = new UrlImportDialog();
        if (dialog.ShowDialog(FindForm()) == DialogResult.OK && dialog.ImportedUrls.Count > 0)
        {
            int addedCount = 0;
            int selectedCount = 0;

            foreach (var importedUrl in dialog.ImportedUrls)
            {
                // Check if site already exists in the list
                var existingSite = _allSites.FirstOrDefault(s =>
                    s.Url.Equals(importedUrl, StringComparison.OrdinalIgnoreCase));

                if (existingSite == null)
                {
                    // Add new site to the list
                    var newSite = new SiteCollection
                    {
                        Url = importedUrl,
                        Title = GetSiteTitleFromUrl(importedUrl),
                        Template = "Unknown"
                    };
                    _allSites.Add(newSite);
                    addedCount++;
                }
            }

            // Refresh the grid
            ApplyFilter();

            // Select all imported sites
            foreach (DataGridViewRow row in _sitesGrid.Rows)
            {
                if (row.Tag is SiteCollection site)
                {
                    var isMatch = dialog.ImportedUrls.Any(url =>
                        url.Equals(site.Url, StringComparison.OrdinalIgnoreCase));
                    if (isMatch)
                    {
                        row.Cells["Select"].Value = true;
                        selectedCount++;
                    }
                }
            }

            UpdateCountLabel();

            var message = addedCount > 0
                ? $"Added {addedCount} new sites. Selected {selectedCount} sites total."
                : $"Selected {selectedCount} sites from import.";
            SetStatus(message);
        }
    }

    private static string GetSiteTitleFromUrl(string url)
    {
        try
        {
            var uri = new Uri(url);
            var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length > 0)
            {
                // Return last segment, replacing dashes with spaces
                return segments[^1].Replace("-", " ");
            }
            return uri.Host;
        }
        catch
        {
            return url;
        }
    }

    private static string GetSitePath(string url)
    {
        try
        {
            var uri = new Uri(url);
            var path = uri.AbsolutePath;
            return string.IsNullOrEmpty(path) || path == "/" ? "/" : path;
        }
        catch
        {
            return url;
        }
    }

    private void SelectAllButton_Click(object? sender, EventArgs e)
    {
        foreach (DataGridViewRow row in _sitesGrid.Rows)
        {
            row.Cells["Select"].Value = true;
        }
        UpdateCountLabel();
    }

    private void SelectNoneButton_Click(object? sender, EventArgs e)
    {
        foreach (DataGridViewRow row in _sitesGrid.Rows)
        {
            row.Cells["Select"].Value = false;
        }
        UpdateCountLabel();
    }

    private async void CreateTaskButton_Click(object? sender, EventArgs e)
    {
        var selectedSites = GetSelectedSites();
        if (selectedSites.Count == 0)
        {
            MessageBox.Show("Please select at least one site.", "No Selection",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var context = new TaskCreationContext
        {
            Connection = _connection,
            SelectedSites = selectedSites
        };

        await NavigationService!.NavigateToAsync<TaskTypeSelectionScreen>(context);
    }

    private async void DeleteButton_Click(object? sender, EventArgs e)
    {
        var selectedSites = GetSelectedSites();
        if (selectedSites.Count == 0)
        {
            MessageBox.Show("Please select at least one site.", "No Selection",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Show confirmation dialog
        using var confirmDialog = new DeleteSitesConfirmationDialog(selectedSites);
        if (confirmDialog.ShowDialog(FindForm()) != DialogResult.Yes)
        {
            return;
        }

        // Proceed with deletion
        await DeleteSelectedSitesAsync(selectedSites);
    }

    private async Task DeleteSelectedSitesAsync(List<SiteCollection> sites)
    {
        ShowLoading($"Deleting {sites.Count} site(s)...");
        var successCount = 0;
        var failedSites = new List<(SiteCollection site, string error)>();

        try
        {
            // Get cookies for the admin domain
            var adminDomain = _connection.AdminDomain;
            var cookies = _authService.GetStoredCookies(adminDomain);

            if (cookies == null || !cookies.IsValid)
            {
                // Try re-authentication
                var authenticated = await ReauthenticateAsync();
                if (!authenticated)
                {
                    HideLoading();
                    return;
                }
                cookies = _authService.GetStoredCookies(adminDomain);
            }

            if (cookies == null)
            {
                MessageBox.Show("No valid credentials available for admin operations.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                HideLoading();
                return;
            }

            using var spService = new SharePointService(cookies, adminDomain);

            for (int i = 0; i < sites.Count; i++)
            {
                var site = sites[i];
                SetStatus($"Deleting ({i + 1}/{sites.Count}): {site.Title}...");

                var result = await spService.DeleteSiteCollectionAsync(site.Url);

                if (result.IsSuccess)
                {
                    successCount++;
                    // Remove from the local lists
                    _allSites.RemoveAll(s => s.Url.Equals(site.Url, StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    failedSites.Add((site, result.ErrorMessage ?? "Unknown error"));
                }
            }

            // Refresh the grid
            ApplyFilter();

            // Show results
            if (failedSites.Count == 0)
            {
                MessageBox.Show(
                    $"Successfully deleted {successCount} site(s).\n\nThe sites have been moved to the SharePoint Recycle Bin.",
                    "Deletion Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else
            {
                var errorMessage = $"Deleted {successCount} of {sites.Count} site(s).\n\n" +
                                   $"Failed to delete {failedSites.Count} site(s):\n\n";
                foreach (var (site, error) in failedSites.Take(5))
                {
                    errorMessage += $"- {site.Title}: {error}\n";
                }
                if (failedSites.Count > 5)
                {
                    errorMessage += $"\n...and {failedSites.Count - 5} more.";
                }

                MessageBox.Show(errorMessage, "Deletion Partially Failed",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            SetStatus($"Deleted {successCount} of {sites.Count} sites");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Deletion failed: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Deletion failed");
        }
        finally
        {
            HideLoading();
        }
    }

    private void SitesGrid_CurrentCellDirtyStateChanged(object? sender, EventArgs e)
    {
        if (_sitesGrid.IsCurrentCellDirty && _sitesGrid.CurrentCell is DataGridViewCheckBoxCell)
        {
            _sitesGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
    }

    private void SitesGrid_CellValueChanged(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.ColumnIndex == 0 && e.RowIndex >= 0)
        {
            UpdateCountLabel();
        }
    }

    private void SitesGrid_ColumnHeaderMouseClick(object? sender, DataGridViewCellMouseEventArgs e)
    {
        // Skip the checkbox column
        if (e.ColumnIndex == 0) return;

        // Determine sort order
        if (_sortColumnIndex == e.ColumnIndex)
        {
            _sortOrder = _sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
        }
        else
        {
            _sortColumnIndex = e.ColumnIndex;
            _sortOrder = SortOrder.Ascending;
        }

        // Sort the filtered sites list
        var columnName = _sitesGrid.Columns[e.ColumnIndex].Name;

        // For StorageUsed column, sort by bytes value
        if (columnName == "StorageUsed")
        {
            _filteredSites = _sortOrder == SortOrder.Ascending
                ? _filteredSites.OrderBy(s => s.StorageUsedBytes).ToList()
                : _filteredSites.OrderByDescending(s => s.StorageUsedBytes).ToList();
        }
        else
        {
            _filteredSites = columnName switch
            {
                "Title" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.Title).ToList()
                    : _filteredSites.OrderByDescending(s => s.Title).ToList(),
                "Url" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.Url).ToList()
                    : _filteredSites.OrderByDescending(s => s.Url).ToList(),
                "SiteType" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.SiteTypeDescription).ToList()
                    : _filteredSites.OrderByDescending(s => s.SiteTypeDescription).ToList(),
                "Template" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.Template).ToList()
                    : _filteredSites.OrderByDescending(s => s.Template).ToList(),
                "Owner" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.Owner).ToList()
                    : _filteredSites.OrderByDescending(s => s.Owner).ToList(),
                "Files" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.FileCount).ToList()
                    : _filteredSites.OrderByDescending(s => s.FileCount).ToList(),
                "PageViews" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.PageViews).ToList()
                    : _filteredSites.OrderByDescending(s => s.PageViews).ToList(),
                "LastActivity" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.LastActivityDate).ToList()
                    : _filteredSites.OrderByDescending(s => s.LastActivityDate).ToList(),
                "ExternalSharing" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.ExternalSharing).ToList()
                    : _filteredSites.OrderByDescending(s => s.ExternalSharing).ToList(),
                "State" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.State).ToList()
                    : _filteredSites.OrderByDescending(s => s.State).ToList(),
                "Standalone" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => !s.IsGroupConnected && s.ChannelType == 0).ToList()
                    : _filteredSites.OrderByDescending(s => !s.IsGroupConnected && s.ChannelType == 0).ToList(),
                "Group" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.IsGroupConnected).ToList()
                    : _filteredSites.OrderByDescending(s => s.IsGroupConnected).ToList(),
                "Channel" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.ChannelType).ToList()
                    : _filteredSites.OrderByDescending(s => s.ChannelType).ToList(),
                "Hub" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.HubDisplay).ToList()
                    : _filteredSites.OrderByDescending(s => s.HubDisplay).ToList(),
                "Language" => _sortOrder == SortOrder.Ascending
                    ? _filteredSites.OrderBy(s => s.LanguageDisplay).ToList()
                    : _filteredSites.OrderByDescending(s => s.LanguageDisplay).ToList(),
                _ => _filteredSites
            };
        }

        RefreshGrid();

        // Set sort glyph
        _sitesGrid.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = _sortOrder;
    }

    private void SitesGrid_SelectionChanged(object? sender, EventArgs e)
    {
        _exploreButton.Enabled = _sitesGrid.SelectedRows.Count == 1;
    }

    private async void SitesGrid_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        // Don't navigate if clicking on the checkbox column
        if (e.ColumnIndex == 0 || e.RowIndex < 0)
            return;

        await ExploreSelectedSiteAsync();
    }

    private async void ExploreButton_Click(object? sender, EventArgs e)
    {
        await ExploreSelectedSiteAsync();
    }

    private async Task ExploreSelectedSiteAsync()
    {
        if (_sitesGrid.SelectedRows.Count == 0)
            return;

        var site = _sitesGrid.SelectedRows[0].Tag as SiteCollection;
        if (site == null)
            return;

        var context = new SiteExplorerContext
        {
            SiteUrl = site.Url,
            SiteTitle = site.Title,
            Connection = _connection
        };

        await NavigationService!.NavigateToAsync<SiteExplorerScreen>(context);
    }

    private void ExportButton_Click(object? sender, EventArgs e)
    {
        if (_filteredSites.Count == 0)
        {
            MessageBox.Show("No sites to export.", "Export",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        // Use tenant name for cleaner filename, fallback to sanitized connection name
        var safeName = !string.IsNullOrEmpty(_connection.TenantName)
            ? _connection.TenantName
            : string.Join("_", _connection.Name.Split(Path.GetInvalidFileNameChars()));

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = $"SiteCollections_{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
        };

        if (dialog.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var exporter = GetRequiredService<CsvExporter>();
                exporter.ExportSiteCollections(_filteredSites, dialog.FileName);
                SetStatus($"Exported {_filteredSites.Count} sites to {dialog.FileName}");

                var result = MessageBox.Show(
                    "Export completed. Would you like to open the file?",
                    "Export Complete",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = dialog.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

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
    private DataGridView _sitesGrid = null!;
    private ComboBox _filterComboBox = null!;
    private TextBox _searchTextBox = null!;
    private Button _refreshButton = null!;
    private Button _importUrlsButton = null!;
    private Button _createTaskButton = null!;
    private Button _selectAllButton = null!;
    private Button _selectNoneButton = null!;
    private Label _countLabel = null!;

    private Connection _connection = null!;
    private List<SiteCollection> _allSites = [];
    private List<SiteCollection> _filteredSites = [];
    private IAuthenticationService _authService = null!;

    public override string ScreenTitle => $"Site Collections - {_connection?.Name ?? "Unknown"}";

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
            Text = "Refresh",
            Size = new Size(80, 28),
            Margin = new Padding(0, 0, 10, 0)
        };
        _refreshButton.Click += RefreshButton_Click;

        _importUrlsButton = new Button
        {
            Text = "Import URLs",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0)
        };
        _importUrlsButton.Click += ImportUrlsButton_Click;

        toolbarPanel.Controls.AddRange(new Control[]
        {
            filterLabel, _filterComboBox,
            searchLabel, _searchTextBox,
            _refreshButton, _importUrlsButton
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
            Text = "Select All",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 10, 0)
        };
        _selectAllButton.Click += SelectAllButton_Click;

        _selectNoneButton = new Button
        {
            Text = "Select None",
            Size = new Size(100, 28),
            Margin = new Padding(0, 0, 20, 0)
        };
        _selectNoneButton.Click += SelectNoneButton_Click;

        _createTaskButton = new Button
        {
            Text = "Create Task from Selection",
            Size = new Size(180, 28),
            Enabled = false
        };
        _createTaskButton.Click += CreateTaskButton_Click;

        _countLabel = new Label
        {
            AutoSize = true,
            Margin = new Padding(20, 8, 0, 0)
        };

        actionPanel.Controls.AddRange(new Control[]
        {
            _selectAllButton, _selectNoneButton, _createTaskButton, _countLabel
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
        _sitesGrid.Columns.Add("Owner", "Owner");

        _sitesGrid.CellValueChanged += SitesGrid_CellValueChanged;
        _sitesGrid.CurrentCellDirtyStateChanged += SitesGrid_CurrentCellDirtyStateChanged;

        Controls.Add(_sitesGrid);
        Controls.Add(actionPanel);
        Controls.Add(toolbarPanel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        if (parameter is Connection connection)
        {
            _connection = connection;
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

        _filteredSites = _allSites.Where(site =>
        {
            // Apply type filter
            if (filterIndex == 1 && site.SiteType == SiteType.OneDrive)
                return false;
            if (filterIndex == 2 && site.SiteType != SiteType.OneDrive)
                return false;

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
        _sitesGrid.Rows.Clear();

        foreach (var site in _filteredSites)
        {
            var rowIndex = _sitesGrid.Rows.Add(
                false,
                site.Title,
                site.Url,
                site.SiteTypeDescription,
                site.Template,
                site.StorageUsedFormatted,
                site.Owner
            );
            _sitesGrid.Rows[rowIndex].Tag = site;
        }

        UpdateCountLabel();
    }

    private void UpdateCountLabel()
    {
        var selectedCount = GetSelectedSites().Count;
        _countLabel.Text = $"{selectedCount} of {_filteredSites.Count} sites selected (Total: {_allSites.Count})";
        _createTaskButton.Enabled = selectedCount > 0;
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
}

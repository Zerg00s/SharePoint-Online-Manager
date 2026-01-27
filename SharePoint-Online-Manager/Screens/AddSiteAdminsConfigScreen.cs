using System.Text.Json;
using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Configuration screen for Add Site Collection Administrators task.
/// </summary>
public class AddSiteAdminsConfigScreen : BaseScreen
{
    private const int MaxAdmins = 5;

    private Label _headerLabel = null!;
    private Label _subHeaderLabel = null!;
    private TextBox _taskNameTextBox = null!;
    private TextBox[] _adminTextBoxes = null!;
    private Button[] _clearButtons = null!;
    private Button[] _searchButtons = null!;
    private Button _createButton = null!;
    private Button _cancelButton = null!;
    private Label _sitesLabel = null!;

    // Store selected users
    private UserSearchResult?[] _selectedUsers = new UserSearchResult?[MaxAdmins];

    private TaskCreationContext _context = null!;
    private ITaskService _taskService = null!;
    private ISharePointService? _sharePointService;
    private IAuthenticationService _authService = null!;
    private string? _siteUrl;

    public override string ScreenTitle => "Add Site Collection Administrators";

    protected override void OnInitialize()
    {
        _taskService = GetRequiredService<ITaskService>();
        _authService = GetRequiredService<IAuthenticationService>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();
        AutoScroll = true;

        // Header panel
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 60,
            Padding = new Padding(0)
        };

        _headerLabel = new Label
        {
            Text = "Add Site Collection Administrators",
            Font = new Font(Font.FontFamily, 14, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(0, 5)
        };

        _subHeaderLabel = new Label
        {
            Text = "Add up to 5 administrators to selected sites.",
            AutoSize = true,
            Location = new Point(0, 32),
            ForeColor = SystemColors.GrayText
        };

        headerPanel.Controls.Add(_headerLabel);
        headerPanel.Controls.Add(_subHeaderLabel);

        // Content panel
        var contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 10, 0, 0)
        };

        int yOffset = 0;

        // Task name
        var nameLabel = new Label
        {
            Text = "Task Name:",
            Location = new Point(0, yOffset),
            AutoSize = true
        };
        contentPanel.Controls.Add(nameLabel);
        yOffset += 22;

        _taskNameTextBox = new TextBox
        {
            Location = new Point(0, yOffset),
            Size = new Size(450, 23),
            Text = $"Add Site Admins - {DateTime.Now:yyyy-MM-dd HH:mm}"
        };
        contentPanel.Controls.Add(_taskNameTextBox);
        yOffset += 40;

        // Administrator fields - using plain TextBoxes
        _adminTextBoxes = new TextBox[MaxAdmins];
        _clearButtons = new Button[MaxAdmins];
        _searchButtons = new Button[MaxAdmins];

        for (int i = 0; i < MaxAdmins; i++)
        {
            var adminLabel = new Label
            {
                Text = $"Administrator {i + 1}:",
                Location = new Point(0, yOffset),
                AutoSize = true
            };
            contentPanel.Controls.Add(adminLabel);
            yOffset += 22;

            var textBox = new TextBox
            {
                Location = new Point(0, yOffset),
                Size = new Size(300, 23),
                PlaceholderText = "Type to search for user...",
                Tag = i
            };
            textBox.TextChanged += AdminTextBox_TextChanged;
            _adminTextBoxes[i] = textBox;
            contentPanel.Controls.Add(textBox);

            var searchBtn = new Button
            {
                Text = "Search",
                Location = new Point(310, yOffset),
                Size = new Size(60, 23),
                FlatStyle = FlatStyle.Flat,
                Tag = i
            };
            searchBtn.Click += SearchButton_Click;
            _searchButtons[i] = searchBtn;
            contentPanel.Controls.Add(searchBtn);

            var clearBtn = new Button
            {
                Text = "Clear",
                Location = new Point(380, yOffset),
                Size = new Size(60, 23),
                FlatStyle = FlatStyle.Flat,
                Tag = i
            };
            clearBtn.FlatAppearance.BorderColor = SystemColors.ControlDark;
            clearBtn.Click += ClearButton_Click;
            _clearButtons[i] = clearBtn;
            contentPanel.Controls.Add(clearBtn);

            yOffset += 35;
        }

        yOffset += 10;

        // Sites count label
        _sitesLabel = new Label
        {
            Text = "Sites to update: 0 selected",
            Location = new Point(0, yOffset),
            AutoSize = true,
            ForeColor = SystemColors.GrayText
        };
        contentPanel.Controls.Add(_sitesLabel);
        yOffset += 35;

        // Buttons panel
        var buttonPanel = new FlowLayoutPanel
        {
            Location = new Point(0, yOffset),
            Size = new Size(500, 40),
            FlowDirection = FlowDirection.LeftToRight
        };

        _cancelButton = new Button
        {
            Text = "Cancel",
            Size = new Size(90, 28),
            Margin = new Padding(0, 0, 10, 0),
            FlatStyle = FlatStyle.Flat
        };
        _cancelButton.FlatAppearance.BorderColor = SystemColors.ControlDark;
        _cancelButton.Click += CancelButton_Click;

        _createButton = new Button
        {
            Text = "Create Task",
            Size = new Size(100, 28),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White,
            Enabled = false
        };
        _createButton.FlatAppearance.BorderSize = 0;
        _createButton.Click += CreateButton_Click;

        buttonPanel.Controls.Add(_cancelButton);
        buttonPanel.Controls.Add(_createButton);
        contentPanel.Controls.Add(buttonPanel);

        Controls.Add(contentPanel);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    private void AdminTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (sender is TextBox textBox && textBox.Tag is int index)
        {
            // Clear selection if user edits the text
            if (_selectedUsers[index] != null)
            {
                var selectedText = _selectedUsers[index]!.ToString();
                if (textBox.Text != selectedText)
                {
                    _selectedUsers[index] = null;
                    textBox.ForeColor = SystemColors.ControlText;
                }
            }
            UpdateCreateButtonState();
        }
    }

    private async void SearchButton_Click(object? sender, EventArgs e)
    {
        if (sender is not Button button || button.Tag is not int index)
            return;

        var query = _adminTextBoxes[index].Text.Trim();
        if (query.Length < 2)
        {
            MessageBox.Show("Please enter at least 2 characters to search.", "Search",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (_sharePointService == null || string.IsNullOrEmpty(_siteUrl))
        {
            MessageBox.Show("SharePoint service not available. Please ensure you are authenticated.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        button.Enabled = false;
        button.Text = "...";

        try
        {
            var result = await _sharePointService.SearchUsersAsync(_siteUrl, query);

            if (result.IsSuccess && result.Data != null && result.Data.Count > 0)
            {
                // Show selection dialog
                using var dialog = new UserSelectionDialog(result.Data, query);
                if (dialog.ShowDialog(FindForm()) == DialogResult.OK && dialog.SelectedUser != null)
                {
                    _selectedUsers[index] = dialog.SelectedUser;
                    _adminTextBoxes[index].Text = dialog.SelectedUser.ToString();
                    _adminTextBoxes[index].ForeColor = Color.DarkGreen;
                    UpdateCreateButtonState();
                }
            }
            else if (result.IsSuccess)
            {
                MessageBox.Show($"No users found matching '{query}'.", "Search Results",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show($"Search failed: {result.ErrorMessage}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Search error: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            button.Enabled = true;
            button.Text = "Search";
        }
    }

    private void ClearButton_Click(object? sender, EventArgs e)
    {
        if (sender is Button button && button.Tag is int index)
        {
            _adminTextBoxes[index].Text = string.Empty;
            _adminTextBoxes[index].ForeColor = SystemColors.ControlText;
            _selectedUsers[index] = null;
            UpdateCreateButtonState();
        }
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        if (parameter is TaskCreationContext context)
        {
            _context = context;
            _sitesLabel.Text = $"Sites to update: {_context.SelectedSites.Count} selected";

            // Get SharePoint service for people picker search
            var cookies = _authService.GetStoredCookies(_context.Connection.AdminDomain);

            if (cookies != null && cookies.IsValid)
            {
                // Use AdminDomain since we're calling admin APIs
                _sharePointService = new SharePointService(cookies, _context.Connection.AdminDomain);
                _siteUrl = _context.Connection.AdminUrl;
                System.Diagnostics.Debug.WriteLine($"[AddSiteAdminsConfig] Using siteUrl for search: {_siteUrl}, domain: {_context.Connection.AdminDomain}");
            }
            else
            {
                SetStatus("Authentication required for user search");
            }
        }
        else
        {
            MessageBox.Show("No context provided.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            await NavigationService!.GoBackAsync();
        }
    }

    public override Task<bool> OnNavigatingFromAsync()
    {
        _sharePointService?.Dispose();
        return Task.FromResult(true);
    }

    private void UpdateCreateButtonState()
    {
        // Enable create button if at least one admin is selected
        var hasSelection = _selectedUsers.Any(u => u != null);
        _createButton.Enabled = hasSelection;
    }

    private async void CancelButton_Click(object? sender, EventArgs e)
    {
        await NavigationService!.GoBackAsync();
    }

    private async void CreateButton_Click(object? sender, EventArgs e)
    {
        // Gather selected administrators
        var admins = _selectedUsers.Where(u => u != null).Select(u => u!).ToList();

        if (admins.Count == 0)
        {
            MessageBox.Show("Please select at least one administrator using the Search button.", "Validation Error",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Create configuration
        var config = new AddSiteAdminsConfiguration
        {
            Administrators = admins
        };

        // Create task
        var task = new TaskDefinition
        {
            Name = _taskNameTextBox.Text.Trim(),
            Type = TaskType.AddSiteCollectionAdmins,
            ConnectionId = _context.Connection.Id,
            TargetSiteUrls = _context.SelectedSites.Select(s => s.Url).ToList(),
            ConfigurationJson = JsonSerializer.Serialize(config),
            Status = Models.TaskStatus.Pending
        };

        await _taskService.SaveTaskAsync(task);
        SetStatus($"Task '{task.Name}' created with {admins.Count} administrator(s)");

        // Navigate to detail screen
        await NavigationService!.NavigateToAsync<AddSiteAdminsDetailScreen>(task);
    }
}

/// <summary>
/// Dialog for selecting a user from search results.
/// </summary>
public class UserSelectionDialog : Form
{
    private ListBox _listBox = null!;
    public UserSearchResult? SelectedUser { get; private set; }

    public UserSelectionDialog(List<UserSearchResult> users, string query)
    {
        Text = $"Select User - '{query}'";
        Size = new Size(450, 300);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;

        var label = new Label
        {
            Text = $"Found {users.Count} user(s). Double-click to select:",
            Location = new Point(15, 15),
            AutoSize = true
        };

        _listBox = new ListBox
        {
            Location = new Point(15, 40),
            Size = new Size(400, 170),
            DisplayMember = "ToString"
        };

        foreach (var user in users)
        {
            _listBox.Items.Add(user);
        }

        _listBox.DoubleClick += (s, e) =>
        {
            if (_listBox.SelectedItem is UserSearchResult user)
            {
                SelectedUser = user;
                DialogResult = DialogResult.OK;
                Close();
            }
        };

        var selectButton = new Button
        {
            Text = "Select",
            Location = new Point(260, 220),
            Size = new Size(75, 28),
            DialogResult = DialogResult.OK
        };
        selectButton.Click += (s, e) =>
        {
            SelectedUser = _listBox.SelectedItem as UserSearchResult;
        };

        var cancelButton = new Button
        {
            Text = "Cancel",
            Location = new Point(340, 220),
            Size = new Size(75, 28),
            DialogResult = DialogResult.Cancel
        };

        AcceptButton = selectButton;
        CancelButton = cancelButton;

        Controls.AddRange([label, _listBox, selectButton, cancelButton]);
    }
}

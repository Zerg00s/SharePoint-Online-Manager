using SharePointOnlineManager.Models;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Controls;

/// <summary>
/// A reusable people picker control with autocomplete search.
/// </summary>
public class PeoplePickerControl : UserControl
{
    private TextBox _textBox = null!;
    private ListBox _resultsListBox = null!;
    private System.Windows.Forms.Timer _debounceTimer = null!;
    private List<UserSearchResult> _searchResults = [];
    private bool _isSelecting;
    private Form? _dropdownForm;

    /// <summary>
    /// The currently selected user.
    /// </summary>
    public UserSearchResult? SelectedUser { get; private set; }

    /// <summary>
    /// The email of the selected user.
    /// </summary>
    public string SelectedUserEmail => SelectedUser?.Email ?? string.Empty;

    /// <summary>
    /// The display name of the selected user.
    /// </summary>
    public string SelectedUserDisplayName => SelectedUser?.DisplayName ?? string.Empty;

    /// <summary>
    /// The login name of the selected user (for API calls).
    /// </summary>
    public string SelectedUserLoginName => SelectedUser?.LoginName ?? string.Empty;

    /// <summary>
    /// Indicates whether a user is currently selected.
    /// </summary>
    public bool HasSelection => SelectedUser != null;

    /// <summary>
    /// The SharePoint service to use for user search.
    /// </summary>
    public ISharePointService? SharePointService { get; set; }

    /// <summary>
    /// The site URL to use as search context.
    /// </summary>
    public string? SiteUrl { get; set; }

    /// <summary>
    /// Occurs when a user is selected.
    /// </summary>
    public event EventHandler<UserSearchResult>? UserSelected;

    /// <summary>
    /// Occurs when the selection is cleared.
    /// </summary>
    public event EventHandler? UserCleared;

    public PeoplePickerControl()
    {
        System.Diagnostics.Debug.WriteLine($"[PeoplePicker] Constructor called");
        InitializeUI();
        System.Diagnostics.Debug.WriteLine($"[PeoplePicker] InitializeUI completed");
    }

    private void InitializeUI()
    {
        SuspendLayout();

        Height = 23;
        Width = 300;
        MinimumSize = new Size(200, 23);

        _textBox = new TextBox
        {
            Location = new Point(0, 0),
            Size = new Size(300, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.LightYellow  // DEBUG: make it visible
        };
        _textBox.TextChanged += TextBox_TextChanged;
        _textBox.KeyDown += TextBox_KeyDown;
        _textBox.Leave += TextBox_Leave;

        _debounceTimer = new System.Windows.Forms.Timer
        {
            Interval = 300
        };
        _debounceTimer.Tick += DebounceTimer_Tick;

        Controls.Add(_textBox);

        ResumeLayout(false);
        PerformLayout();

        System.Diagnostics.Debug.WriteLine($"[PeoplePicker] InitializeUI - TextBox size: {_textBox.Size}, Parent size: {Size}");
    }

    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        if (_textBox != null)
        {
            _textBox.Width = Width;
            System.Diagnostics.Debug.WriteLine($"[PeoplePicker] OnSizeChanged - new width: {Width}");
        }
    }

    protected override void OnVisibleChanged(EventArgs e)
    {
        base.OnVisibleChanged(e);
        System.Diagnostics.Debug.WriteLine($"[PeoplePicker] OnVisibleChanged - Visible: {Visible}, TextBox: {_textBox?.Visible}");
    }

    private void TextBox_TextChanged(object? sender, EventArgs e)
    {
        System.Diagnostics.Debug.WriteLine($"[PeoplePicker] TextChanged: '{_textBox.Text}', isSelecting={_isSelecting}");

        if (_isSelecting)
            return;

        // Clear selection when text changes
        if (SelectedUser != null && _textBox.Text != SelectedUser.ToString())
        {
            SelectedUser = null;
            UserCleared?.Invoke(this, EventArgs.Empty);
        }

        // Restart debounce timer
        _debounceTimer.Stop();
        _debounceTimer.Start();
        System.Diagnostics.Debug.WriteLine($"[PeoplePicker] Debounce timer started");
    }

    private async void DebounceTimer_Tick(object? sender, EventArgs e)
    {
        _debounceTimer.Stop();
        System.Diagnostics.Debug.WriteLine($"[PeoplePicker] Debounce timer tick");

        var query = _textBox.Text.Trim();
        if (query.Length < 2)
        {
            System.Diagnostics.Debug.WriteLine($"[PeoplePicker] Query too short: '{query}'");
            HideDropdown();
            return;
        }

        await SearchUsersAsync(query);
    }

    private async Task SearchUsersAsync(string query)
    {
        System.Diagnostics.Debug.WriteLine($"[PeoplePicker] SearchUsersAsync: query='{query}', Service={SharePointService != null}, SiteUrl='{SiteUrl}'");

        if (SharePointService == null || string.IsNullOrEmpty(SiteUrl))
        {
            System.Diagnostics.Debug.WriteLine($"[PeoplePicker] Aborting search - service or URL is null");
            return;
        }

        try
        {
            var result = await SharePointService.SearchUsersAsync(SiteUrl, query);

            if (result.IsSuccess && result.Data != null && result.Data.Count > 0)
            {
                _searchResults = result.Data;
                ShowDropdown();
            }
            else
            {
                HideDropdown();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PeoplePicker] Search error: {ex.Message}");
            HideDropdown();
        }
    }

    private void ShowDropdown()
    {
        if (_dropdownForm == null)
        {
            _dropdownForm = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.Manual,
                TopMost = true
            };

            _resultsListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false
            };
            _resultsListBox.Click += ResultsListBox_Click;
            _resultsListBox.KeyDown += ResultsListBox_KeyDown;

            _dropdownForm.Controls.Add(_resultsListBox);
            _dropdownForm.Deactivate += DropdownForm_Deactivate;
        }

        // Populate results
        _resultsListBox.Items.Clear();
        foreach (var user in _searchResults)
        {
            _resultsListBox.Items.Add(user);
        }

        // Position the dropdown
        var screenPoint = _textBox.PointToScreen(new Point(0, _textBox.Height));
        _dropdownForm.Location = screenPoint;
        _dropdownForm.Width = _textBox.Width;
        _dropdownForm.Height = Math.Min(_searchResults.Count * 20 + 4, 200);

        if (!_dropdownForm.Visible)
        {
            _dropdownForm.Show(FindForm());
        }

        _resultsListBox.Focus();
    }

    private void HideDropdown()
    {
        _dropdownForm?.Hide();
    }

    private void DropdownForm_Deactivate(object? sender, EventArgs e)
    {
        // Small delay to allow click to process
        BeginInvoke(() => HideDropdown());
    }

    private void ResultsListBox_Click(object? sender, EventArgs e)
    {
        SelectCurrentItem();
    }

    private void ResultsListBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter)
        {
            SelectCurrentItem();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Escape)
        {
            HideDropdown();
            _textBox.Focus();
            e.Handled = true;
        }
    }

    private void SelectCurrentItem()
    {
        if (_resultsListBox.SelectedItem is UserSearchResult user)
        {
            SelectUser(user);
        }
    }

    private void SelectUser(UserSearchResult user)
    {
        _isSelecting = true;
        SelectedUser = user;
        _textBox.Text = user.ToString();
        _textBox.ForeColor = Color.DarkGreen;
        HideDropdown();
        _isSelecting = false;

        UserSelected?.Invoke(this, user);
        _textBox.Focus();
    }

    private void TextBox_KeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Down && _dropdownForm?.Visible == true)
        {
            _resultsListBox.Focus();
            if (_resultsListBox.Items.Count > 0)
            {
                _resultsListBox.SelectedIndex = 0;
            }
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Escape)
        {
            HideDropdown();
            e.Handled = true;
        }
    }

    private void TextBox_Leave(object? sender, EventArgs e)
    {
        // Don't hide if focus moved to dropdown
        if (_dropdownForm?.ContainsFocus == true)
            return;

        BeginInvoke(() =>
        {
            if (!_textBox.Focused && _dropdownForm?.ContainsFocus != true)
            {
                HideDropdown();
            }
        });
    }

    /// <summary>
    /// Clears the current selection.
    /// </summary>
    public void Clear()
    {
        _isSelecting = true;
        SelectedUser = null;
        _textBox.Text = string.Empty;
        _textBox.ForeColor = SystemColors.ControlText;
        _isSelecting = false;
        HideDropdown();
        UserCleared?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Sets the selected user programmatically.
    /// </summary>
    public void SetUser(UserSearchResult user)
    {
        SelectUser(user);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _debounceTimer?.Dispose();
            _dropdownForm?.Dispose();
        }
        base.Dispose(disposing);
    }
}

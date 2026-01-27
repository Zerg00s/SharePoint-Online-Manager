using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Screens;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Forms;

/// <summary>
/// Main application shell form for SharePoint Online Manager.
/// Acts as a container for screen navigation.
/// </summary>
public partial class MainForm : Form
{
    private readonly IServiceProvider _serviceProvider;
    private INavigationService _navigationService = null!;

    // Shell UI Controls
    private MenuStrip _menuStrip = null!;
    private ToolStrip _toolStrip = null!;
    private Panel _contentPanel = null!;
    private StatusStrip _statusStrip = null!;
    private ToolStripStatusLabel _statusLabel = null!;
    private ToolStripProgressBar _progressBar = null!;
    private ToolStripButton _backButton = null!;
    private ToolStripButton _homeButton = null!;
    private ToolStripLabel _titleLabel = null!;

    public MainForm(IServiceProvider serviceProvider)
    {
        _serviceProvider = serviceProvider;
        InitializeComponent();
        InitializeShellUI();
    }

    private void InitializeShellUI()
    {
        SuspendLayout();

        // Menu Strip
        _menuStrip = new MenuStrip();

        var fileMenu = new ToolStripMenuItem("&File");
        fileMenu.DropDownItems.Add("E&xit", null, (s, e) => Close());

        var viewMenu = new ToolStripMenuItem("&View");
        viewMenu.DropDownItems.Add("&Home", null, async (s, e) => await _navigationService.NavigateToHomeAsync());
        viewMenu.DropDownItems.Add(new ToolStripSeparator());
        viewMenu.DropDownItems.Add("&Tasks", null, async (s, e) => await NavigateToTaskListAsync());

        var helpMenu = new ToolStripMenuItem("&Help");
        helpMenu.DropDownItems.Add("&About", null, (s, e) => ShowAbout());

        _menuStrip.Items.AddRange(new ToolStripItem[] { fileMenu, viewMenu, helpMenu });
        _menuStrip.Location = new Point(0, 0);
        _menuStrip.Name = "menuStrip";

        // Tool Strip
        _toolStrip = new ToolStrip();

        _backButton = new ToolStripButton
        {
            Text = "\u2190 Back", // Left arrow
            DisplayStyle = ToolStripItemDisplayStyle.Text,
            Enabled = false,
            ToolTipText = "Go back to previous screen",
            Font = new Font("Segoe UI", 9F)
        };
        _backButton.Click += async (s, e) => await _navigationService.GoBackAsync();

        _homeButton = new ToolStripButton
        {
            Text = "\U0001F3E0 Home", // House emoji
            DisplayStyle = ToolStripItemDisplayStyle.Text,
            ToolTipText = "Go to Connections",
            Font = new Font("Segoe UI", 9F)
        };
        _homeButton.Click += async (s, e) => await _navigationService.NavigateToHomeAsync();

        var tasksButton = new ToolStripButton
        {
            Text = "\U0001F4CB Tasks", // Clipboard emoji
            DisplayStyle = ToolStripItemDisplayStyle.Text,
            ToolTipText = "View all tasks",
            Font = new Font("Segoe UI", 9F)
        };
        tasksButton.Click += async (s, e) => await NavigateToTaskListAsync();

        var separator = new ToolStripSeparator();

        _titleLabel = new ToolStripLabel
        {
            Text = "SharePoint Online Manager",
            Font = new Font(Font.FontFamily, 10F, FontStyle.Bold)
        };

        _toolStrip.Items.AddRange(new ToolStripItem[] { _backButton, _homeButton, tasksButton, separator, _titleLabel });
        _toolStrip.Location = new Point(0, 24);
        _toolStrip.Name = "toolStrip";

        // Content Panel
        _contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Name = "contentPanel",
            Padding = new Padding(8)
        };

        // Status Strip
        _statusStrip = new StatusStrip();

        _statusLabel = new ToolStripStatusLabel
        {
            Text = "Ready",
            Spring = true,
            TextAlign = ContentAlignment.MiddleLeft
        };

        _progressBar = new ToolStripProgressBar
        {
            Style = ProgressBarStyle.Marquee,
            Visible = false,
            Width = 150
        };

        _statusStrip.Items.AddRange(new ToolStripItem[] { _statusLabel, _progressBar });
        _statusStrip.Location = new Point(0, 539);
        _statusStrip.Name = "statusStrip";

        // Add controls in order (order matters for docking)
        Controls.Add(_contentPanel);
        Controls.Add(_toolStrip);
        Controls.Add(_menuStrip);
        Controls.Add(_statusStrip);

        MainMenuStrip = _menuStrip;

        ResumeLayout(true);

        // Initialize navigation service
        _navigationService = new NavigationService(
            _contentPanel,
            _serviceProvider,
            SetStatus,
            ShowLoading,
            HideLoading,
            SetTitle,
            SetBackButtonEnabled
        );

        // Navigate to home screen on load
        Load += async (s, e) => await NavigateToHomeAsync();
    }

    private async Task NavigateToHomeAsync()
    {
        await _navigationService.NavigateToAsync<HomeScreen>();
    }

    private async Task NavigateToTaskListAsync()
    {
        await _navigationService.NavigateToAsync<TaskListScreen>();
    }

    private void SetStatus(string message)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetStatus(message));
            return;
        }
        _statusLabel.Text = message;
    }

    private void ShowLoading(string message)
    {
        if (InvokeRequired)
        {
            Invoke(() => ShowLoading(message));
            return;
        }
        _statusLabel.Text = message;
        _progressBar.Visible = true;
    }

    private void HideLoading()
    {
        if (InvokeRequired)
        {
            Invoke(HideLoading);
            return;
        }
        _progressBar.Visible = false;
        _statusLabel.Text = "Ready";
    }

    private void SetTitle(string title)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetTitle(title));
            return;
        }
        _titleLabel.Text = title;
        Text = $"SharePoint Online Manager - {title}";
    }

    private void SetBackButtonEnabled(bool enabled)
    {
        if (InvokeRequired)
        {
            Invoke(() => SetBackButtonEnabled(enabled));
            return;
        }
        _backButton.Enabled = enabled;
    }

    private static void ShowAbout()
    {
        MessageBox.Show(
            "SharePoint Online Manager\n\n" +
            "Version 1.0\n\n" +
            "A tool for managing SharePoint Online sites and running reports.\n\n" +
            "Developed by Denis Molodtsov\n" +
            "\u00A9 2026 All Rights Reserved",
            "About SharePoint Online Manager",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }
}

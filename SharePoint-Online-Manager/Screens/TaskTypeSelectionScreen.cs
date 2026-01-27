using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for selecting a task type before creating a new task.
/// </summary>
public class TaskTypeSelectionScreen : BaseScreen
{
    private FlowLayoutPanel _taskTypesPanel = null!;
    private Label _headerLabel = null!;
    private TaskCreationContext _context = null!;

    public override string ScreenTitle => "Select Task Type";

    protected override void OnInitialize()
    {
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();

        // Header panel
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 60,
            Padding = new Padding(10)
        };

        _headerLabel = new Label
        {
            Text = "Select Task Type",
            Font = new Font(Font.FontFamily, 14, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(10, 10)
        };

        var subHeaderLabel = new Label
        {
            Name = "SubHeaderLabel",
            Text = "0 sites selected",
            AutoSize = true,
            Location = new Point(10, 35),
            ForeColor = SystemColors.GrayText
        };

        headerPanel.Controls.Add(_headerLabel);
        headerPanel.Controls.Add(subHeaderLabel);

        // Task types panel (scrollable)
        _taskTypesPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            Padding = new Padding(10)
        };

        Controls.Add(_taskTypesPanel);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    public override Task OnNavigatedToAsync(object? parameter = null)
    {
        if (parameter is TaskCreationContext context)
        {
            _context = context;
            var subHeader = Controls.Find("SubHeaderLabel", true).FirstOrDefault() as Label;
            if (subHeader != null)
            {
                subHeader.Text = $"{_context.SelectedSites.Count} site(s) selected";
            }
        }
        else
        {
            MessageBox.Show("No context provided.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            _ = NavigationService!.GoBackAsync();
            return Task.CompletedTask;
        }

        PopulateTaskTypes();
        return Task.CompletedTask;
    }

    private void PopulateTaskTypes()
    {
        _taskTypesPanel.Controls.Clear();

        foreach (TaskType taskType in Enum.GetValues<TaskType>())
        {
            var card = CreateTaskTypeCard(taskType);
            _taskTypesPanel.Controls.Add(card);
        }
    }

    private Panel CreateTaskTypeCard(TaskType taskType)
    {
        var card = new Panel
        {
            Size = new Size(_taskTypesPanel.ClientSize.Width - 40, 80),
            Margin = new Padding(0, 0, 0, 10),
            BackColor = SystemColors.Window,
            BorderStyle = BorderStyle.FixedSingle,
            Cursor = Cursors.Hand,
            Tag = taskType
        };

        var iconLabel = new Label
        {
            Text = GetTaskTypeIcon(taskType),
            Font = new Font("Segoe UI", 20),
            Location = new Point(15, 15),
            AutoSize = true
        };

        var nameLabel = new Label
        {
            Text = taskType.GetDisplayName(),
            Font = new Font(Font.FontFamily, 11, FontStyle.Bold),
            Location = new Point(75, 15),
            AutoSize = true
        };

        var descLabel = new Label
        {
            Text = taskType.GetDescription(),
            Location = new Point(75, 40),
            AutoSize = true,
            ForeColor = SystemColors.GrayText
        };

        card.Controls.Add(iconLabel);
        card.Controls.Add(nameLabel);
        card.Controls.Add(descLabel);

        // Hover effects
        card.MouseEnter += (s, e) => card.BackColor = SystemColors.ControlLight;
        card.MouseLeave += (s, e) => card.BackColor = SystemColors.Window;

        // Apply hover effects to child controls too
        foreach (Control child in card.Controls)
        {
            child.MouseEnter += (s, e) => card.BackColor = SystemColors.ControlLight;
            child.MouseLeave += (s, e) => card.BackColor = SystemColors.Window;
            child.Click += (s, e) => OnTaskTypeSelected(taskType);
        }

        card.Click += (s, e) => OnTaskTypeSelected(taskType);

        // Handle resize to adjust card width
        _taskTypesPanel.Resize += (s, e) =>
        {
            card.Width = _taskTypesPanel.ClientSize.Width - 40;
        };

        return card;
    }

    private static string GetTaskTypeIcon(TaskType taskType) => taskType switch
    {
        TaskType.ListsReport => "\U0001F4CB", // clipboard emoji
        TaskType.ListCompare => "\U0001F504", // arrows clockwise emoji (compare)
        TaskType.DocumentReport => "\U0001F4C1", // file folder emoji (documents)
        TaskType.PermissionReport => "\U0001F512", // lock emoji (permissions)
        TaskType.SetSiteState => "\u2699", // gear emoji (settings)
        TaskType.AddSiteCollectionAdmins => "\U0001F464", // bust in silhouette emoji (user/admin)
        TaskType.RemoveSiteCollectionAdmins => "\U0001F6AB", // no entry sign emoji (remove)
        _ => "\U0001F4C4" // page emoji
    };

    private async void OnTaskTypeSelected(TaskType taskType)
    {
        System.Diagnostics.Debug.WriteLine($"[TaskTypeSelection] OnTaskTypeSelected: {taskType}");

        // AddSiteCollectionAdmins and RemoveSiteCollectionAdmins have their own configuration screens
        if (taskType == TaskType.AddSiteCollectionAdmins)
        {
            await NavigationService!.NavigateToAsync<AddSiteAdminsConfigScreen>(_context);
            return;
        }

        if (taskType == TaskType.RemoveSiteCollectionAdmins)
        {
            await NavigationService!.NavigateToAsync<RemoveSiteAdminsConfigScreen>(_context);
            return;
        }

        using var dialog = new CreateTaskDialog(_context.SelectedSites.Count, taskType);
        if (dialog.ShowDialog(FindForm()) == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.TaskName))
        {
            var taskService = GetRequiredService<ITaskService>();

            var task = new TaskDefinition
            {
                Name = dialog.TaskName,
                Type = taskType,
                ConnectionId = _context.Connection.Id,
                TargetSiteUrls = _context.SelectedSites.Select(s => s.Url).ToList(),
                Status = Models.TaskStatus.Pending
            };

            // Save state configuration for SetSiteState task
            if (taskType == TaskType.SetSiteState && dialog.SelectedState != null)
            {
                var config = new { TargetState = dialog.SelectedState };
                task.ConfigurationJson = System.Text.Json.JsonSerializer.Serialize(config);
            }

            await taskService.SaveTaskAsync(task);

            SetStatus($"Task '{task.Name}' created with {_context.SelectedSites.Count} sites");

            // Navigate to appropriate detail screen based on task type
            if (taskType == TaskType.DocumentReport)
            {
                await NavigationService!.NavigateToAsync<DocumentReportDetailScreen>(task);
            }
            else if (taskType == TaskType.PermissionReport)
            {
                await NavigationService!.NavigateToAsync<PermissionReportDetailScreen>(task);
            }
            else if (taskType == TaskType.SetSiteState)
            {
                await NavigationService!.NavigateToAsync<SetSiteStateDetailScreen>(task);
            }
            else
            {
                await NavigationService!.NavigateToAsync<TaskDetailScreen>(task);
            }
        }
    }
}

/// <summary>
/// Dialog for creating a new task.
/// </summary>
public class CreateTaskDialog : Form
{
    private TextBox _nameTextBox = null!;
    private ComboBox? _stateComboBox;

    public string TaskName { get; private set; } = string.Empty;
    public string? SelectedState { get; private set; }

    public CreateTaskDialog(int siteCount, TaskType taskType)
    {
        InitializeUI(siteCount, taskType);
    }

    private void InitializeUI(int siteCount, TaskType taskType)
    {
        Text = "Create Task";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;

        var infoLabel = new Label
        {
            Text = $"Creating {taskType.GetDisplayName()} task for {siteCount} site(s)",
            Location = new Point(15, 15),
            AutoSize = true
        };

        var nameLabel = new Label
        {
            Text = "Task Name:",
            Location = new Point(15, 50),
            AutoSize = true
        };

        _nameTextBox = new TextBox
        {
            Location = new Point(15, 70),
            Size = new Size(350, 23),
            Text = $"{taskType.GetDisplayName()} - {DateTime.Now:yyyy-MM-dd HH:mm}"
        };

        int buttonY = 105;

        // Add state dropdown for SetSiteState task type
        if (taskType == TaskType.SetSiteState)
        {
            var stateLabel = new Label
            {
                Text = "Target State:",
                Location = new Point(15, 105),
                AutoSize = true
            };

            _stateComboBox = new ComboBox
            {
                Location = new Point(15, 125),
                Size = new Size(350, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _stateComboBox.Items.AddRange(["Unlock (Active)", "Read Only", "No Access (Restricted)"]);
            _stateComboBox.SelectedIndex = 0;

            Controls.Add(stateLabel);
            Controls.Add(_stateComboBox);

            buttonY = 160;
            Size = new Size(400, 235);
        }
        else
        {
            Size = new Size(400, 180);
        }

        var okButton = new Button
        {
            Text = "Create",
            DialogResult = DialogResult.OK,
            Location = new Point(205, buttonY),
            Size = new Size(75, 28)
        };
        okButton.Click += (s, e) =>
        {
            TaskName = _nameTextBox.Text.Trim();
            SelectedState = _stateComboBox?.SelectedItem?.ToString();
        };

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Location = new Point(290, buttonY),
            Size = new Size(75, 28)
        };

        AcceptButton = okButton;
        CancelButton = cancelButton;

        Controls.AddRange(new Control[] { infoLabel, nameLabel, _nameTextBox, okButton, cancelButton });
    }
}

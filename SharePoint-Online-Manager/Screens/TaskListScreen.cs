using SharePointOnlineManager.Models;
using SharePointOnlineManager.Navigation;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager.Screens;

/// <summary>
/// Screen for displaying saved tasks.
/// </summary>
public class TaskListScreen : BaseScreen
{
    private ListView _tasksListView = null!;
    private Button _runButton = null!;
    private Button _viewButton = null!;
    private Button _deleteButton = null!;
    private ITaskService _taskService = null!;
    private IConnectionManager _connectionManager = null!;

    public override string ScreenTitle => "Tasks";

    protected override void OnInitialize()
    {
        _taskService = GetRequiredService<ITaskService>();
        _connectionManager = GetRequiredService<IConnectionManager>();
        InitializeUI();
    }

    private void InitializeUI()
    {
        SuspendLayout();

        // Header panel with buttons
        var headerPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 45,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 5, 0, 5)
        };

        _runButton = new Button
        {
            Text = "Run",
            Size = new Size(80, 32),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false
        };
        _runButton.Click += RunButton_Click;

        _viewButton = new Button
        {
            Text = "View Details",
            Size = new Size(100, 32),
            Margin = new Padding(0, 0, 10, 0),
            Enabled = false
        };
        _viewButton.Click += ViewButton_Click;

        _deleteButton = new Button
        {
            Text = "Delete",
            Size = new Size(80, 32),
            Enabled = false
        };
        _deleteButton.Click += DeleteButton_Click;

        headerPanel.Controls.AddRange(new Control[] { _runButton, _viewButton, _deleteButton });

        // Tasks ListView
        _tasksListView = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            MultiSelect = false
        };
        _tasksListView.Columns.Add("Name", 250);
        _tasksListView.Columns.Add("Type", 120);
        _tasksListView.Columns.Add("Sites", 80);
        _tasksListView.Columns.Add("Status", 100);
        _tasksListView.Columns.Add("Last Run", 150);
        _tasksListView.Columns.Add("Connection", 150);

        _tasksListView.SelectedIndexChanged += TasksListView_SelectedIndexChanged;
        _tasksListView.DoubleClick += TasksListView_DoubleClick;

        Controls.Add(_tasksListView);
        Controls.Add(headerPanel);

        ResumeLayout(true);
    }

    public override async Task OnNavigatedToAsync(object? parameter = null)
    {
        await RefreshTasksAsync();
    }

    private async Task RefreshTasksAsync()
    {
        ShowLoading("Loading tasks...");

        try
        {
            var tasks = await _taskService.GetAllTasksAsync();
            var connections = await _connectionManager.GetAllConnectionsAsync();

            _tasksListView.Items.Clear();

            foreach (var task in tasks)
            {
                var connection = connections.FirstOrDefault(c => c.Id == task.ConnectionId);
                var connectionName = connection?.Name ?? "Unknown";

                var item = new ListViewItem(task.Name)
                {
                    Tag = task
                };
                item.SubItems.Add(task.TypeDescription);
                item.SubItems.Add(task.TotalSites.ToString());
                item.SubItems.Add(task.StatusDescription);
                item.SubItems.Add(task.LastRunAt?.ToString("g") ?? "Never");
                item.SubItems.Add(connectionName);

                // Color code by status
                item.BackColor = task.Status switch
                {
                    Models.TaskStatus.Running => Color.FromArgb(255, 255, 200),
                    Models.TaskStatus.Completed => Color.FromArgb(200, 255, 200),
                    Models.TaskStatus.Failed => Color.FromArgb(255, 200, 200),
                    Models.TaskStatus.Cancelled => Color.FromArgb(255, 220, 180),
                    _ => SystemColors.Window
                };

                _tasksListView.Items.Add(item);
            }

            UpdateButtonStates();
            SetStatus($"Loaded {tasks.Count} tasks");
        }
        finally
        {
            HideLoading();
        }
    }

    private void TasksListView_SelectedIndexChanged(object? sender, EventArgs e)
    {
        UpdateButtonStates();
    }

    private void UpdateButtonStates()
    {
        var hasSelection = _tasksListView.SelectedItems.Count > 0;
        _viewButton.Enabled = hasSelection;
        _deleteButton.Enabled = hasSelection;

        if (hasSelection)
        {
            var task = (TaskDefinition)_tasksListView.SelectedItems[0].Tag;
            _runButton.Enabled = task.Status != Models.TaskStatus.Running;
        }
        else
        {
            _runButton.Enabled = false;
        }
    }

    private async void TasksListView_DoubleClick(object? sender, EventArgs e)
    {
        await ViewSelectedTaskAsync();
    }

    private async void ViewButton_Click(object? sender, EventArgs e)
    {
        await ViewSelectedTaskAsync();
    }

    private async Task ViewSelectedTaskAsync()
    {
        if (_tasksListView.SelectedItems.Count == 0)
            return;

        var task = (TaskDefinition)_tasksListView.SelectedItems[0].Tag;

        // Route to appropriate detail screen based on task type
        if (task.Type == TaskType.ListCompare)
        {
            await NavigationService!.NavigateToAsync<ListCompareDetailScreen>(task);
        }
        else if (task.Type == TaskType.DocumentReport)
        {
            await NavigationService!.NavigateToAsync<DocumentReportDetailScreen>(task);
        }
        else if (task.Type == TaskType.PermissionReport)
        {
            await NavigationService!.NavigateToAsync<PermissionReportDetailScreen>(task);
        }
        else
        {
            await NavigationService!.NavigateToAsync<TaskDetailScreen>(task);
        }
    }

    private async void RunButton_Click(object? sender, EventArgs e)
    {
        if (_tasksListView.SelectedItems.Count == 0)
            return;

        var task = (TaskDefinition)_tasksListView.SelectedItems[0].Tag;

        // Navigate to appropriate detail screen and run
        var execParam = new TaskExecutionParameter
        {
            Task = task,
            ExecuteImmediately = true
        };

        if (task.Type == TaskType.ListCompare)
        {
            await NavigationService!.NavigateToAsync<ListCompareDetailScreen>(execParam);
        }
        else if (task.Type == TaskType.DocumentReport)
        {
            await NavigationService!.NavigateToAsync<DocumentReportDetailScreen>(execParam);
        }
        else if (task.Type == TaskType.PermissionReport)
        {
            await NavigationService!.NavigateToAsync<PermissionReportDetailScreen>(execParam);
        }
        else
        {
            await NavigationService!.NavigateToAsync<TaskDetailScreen>(execParam);
        }
    }

    private async void DeleteButton_Click(object? sender, EventArgs e)
    {
        if (_tasksListView.SelectedItems.Count == 0)
            return;

        var task = (TaskDefinition)_tasksListView.SelectedItems[0].Tag;

        var result = MessageBox.Show(
            $"Are you sure you want to delete the task '{task.Name}'?\n\nThis will also delete all saved results.",
            "Delete Task",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            await _taskService.DeleteTaskAsync(task.Id);
            await RefreshTasksAsync();
            SetStatus($"Task '{task.Name}' deleted");
        }
    }
}

/// <summary>
/// Parameter for navigating to TaskDetailScreen with execution option.
/// </summary>
public class TaskExecutionParameter
{
    public TaskDefinition Task { get; init; } = null!;
    public bool ExecuteImmediately { get; init; }
}

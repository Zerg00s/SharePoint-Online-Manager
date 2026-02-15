using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Interface for task management operations.
/// </summary>
public interface ITaskService
{
    /// <summary>
    /// Gets all task definitions.
    /// </summary>
    Task<List<TaskDefinition>> GetAllTasksAsync();

    /// <summary>
    /// Gets a task by its ID.
    /// </summary>
    Task<TaskDefinition?> GetTaskAsync(Guid id);

    /// <summary>
    /// Saves a task definition.
    /// </summary>
    Task SaveTaskAsync(TaskDefinition task);

    /// <summary>
    /// Deletes a task and its results.
    /// </summary>
    Task DeleteTaskAsync(Guid id);

    /// <summary>
    /// Executes a task and returns the results.
    /// </summary>
    Task<TaskResult> ExecuteTaskAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets all results for a task.
    /// </summary>
    Task<List<TaskResult>> GetTaskResultsAsync(Guid taskId);

    /// <summary>
    /// Gets the most recent result for a task.
    /// </summary>
    Task<TaskResult?> GetLatestTaskResultAsync(Guid taskId);

    /// <summary>
    /// Saves a task result.
    /// </summary>
    Task SaveTaskResultAsync(TaskResult result);

    /// <summary>
    /// Executes a list compare task and returns the results.
    /// </summary>
    Task<ListCompareResult> ExecuteListCompareAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IConnectionManager connectionManager,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets all list compare results for a task.
    /// </summary>
    Task<List<ListCompareResult>> GetListCompareResultsAsync(Guid taskId);

    /// <summary>
    /// Gets the most recent list compare result for a task.
    /// </summary>
    Task<ListCompareResult?> GetLatestListCompareResultAsync(Guid taskId);

    /// <summary>
    /// Saves a list compare result.
    /// </summary>
    Task SaveListCompareResultAsync(ListCompareResult result);

    /// <summary>
    /// Executes a document report task and returns the results.
    /// </summary>
    Task<DocumentReportResult> ExecuteDocumentReportAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets all document report results for a task.
    /// </summary>
    Task<List<DocumentReportResult>> GetDocumentReportResultsAsync(Guid taskId);

    /// <summary>
    /// Gets the most recent document report result for a task.
    /// </summary>
    Task<DocumentReportResult?> GetLatestDocumentReportResultAsync(Guid taskId);

    /// <summary>
    /// Saves a document report result.
    /// </summary>
    Task SaveDocumentReportResultAsync(DocumentReportResult result);

    /// <summary>
    /// Executes a permission report task and returns the results.
    /// </summary>
    Task<PermissionReportResult> ExecutePermissionReportAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets all permission report results for a task.
    /// </summary>
    Task<List<PermissionReportResult>> GetPermissionReportResultsAsync(Guid taskId);

    /// <summary>
    /// Gets the most recent permission report result for a task.
    /// </summary>
    Task<PermissionReportResult?> GetLatestPermissionReportResultAsync(Guid taskId);

    /// <summary>
    /// Saves a permission report result.
    /// </summary>
    Task SavePermissionReportResultAsync(PermissionReportResult result);

    /// <summary>
    /// Executes a navigation settings sync task and returns the results.
    /// </summary>
    Task<NavigationSettingsResult> ExecuteNavigationSettingsSyncAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IConnectionManager connectionManager,
        bool applyMode = false,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets all navigation settings results for a task.
    /// </summary>
    Task<List<NavigationSettingsResult>> GetNavigationSettingsResultsAsync(Guid taskId);

    /// <summary>
    /// Gets the most recent navigation settings result for a task.
    /// </summary>
    Task<NavigationSettingsResult?> GetLatestNavigationSettingsResultAsync(Guid taskId);

    /// <summary>
    /// Saves a navigation settings result.
    /// </summary>
    Task SaveNavigationSettingsResultAsync(NavigationSettingsResult result);

    /// <summary>
    /// Executes a document compare task and returns the results.
    /// </summary>
    /// <param name="continueFromPrevious">If true, skips site pairs that were successfully completed in the previous run.</param>
    /// <param name="reauthCallback">Optional callback invoked on auth failure. Takes (tenantName, tenantDomain), returns fresh cookies or null if dismissed.</param>
    Task<DocumentCompareResult> ExecuteDocumentCompareAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IConnectionManager connectionManager,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default,
        bool continueFromPrevious = false,
        Func<string, string, Task<AuthCookies?>>? reauthCallback = null);

    /// <summary>
    /// Gets all document compare results for a task.
    /// </summary>
    Task<List<DocumentCompareResult>> GetDocumentCompareResultsAsync(Guid taskId);

    /// <summary>
    /// Gets the most recent document compare result for a task.
    /// </summary>
    Task<DocumentCompareResult?> GetLatestDocumentCompareResultAsync(Guid taskId);

    /// <summary>
    /// Saves a document compare result.
    /// </summary>
    Task SaveDocumentCompareResultAsync(DocumentCompareResult result);

    /// <summary>
    /// Executes a site access check task and returns the results.
    /// </summary>
    Task<SiteAccessResult> ExecuteSiteAccessCheckAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IConnectionManager connectionManager,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Gets all site access check results for a task.
    /// </summary>
    Task<List<SiteAccessResult>> GetSiteAccessResultsAsync(Guid taskId);

    /// <summary>
    /// Gets the most recent site access check result for a task.
    /// </summary>
    Task<SiteAccessResult?> GetLatestSiteAccessResultAsync(Guid taskId);

    /// <summary>
    /// Saves a site access check result.
    /// </summary>
    Task SaveSiteAccessResultAsync(SiteAccessResult result);
}

/// <summary>
/// Progress information for task execution.
/// </summary>
public class TaskProgress
{
    public int CurrentSite { get; init; }
    public int TotalSites { get; init; }
    public string CurrentSiteUrl { get; init; } = string.Empty;
    public string Message { get; init; } = string.Empty;
    public int PercentComplete => TotalSites > 0 ? (CurrentSite * 100) / TotalSites : 0;

    /// <summary>
    /// Optional: Completed site result for real-time UI updates (Document Compare).
    /// </summary>
    public SiteDocumentCompareResult? CompletedSiteResult { get; init; }

    /// <summary>
    /// Optional: Completed site pair access result for real-time UI updates (Site Access Check).
    /// </summary>
    public SitePairAccessResult? CompletedAccessPairResult { get; init; }
}

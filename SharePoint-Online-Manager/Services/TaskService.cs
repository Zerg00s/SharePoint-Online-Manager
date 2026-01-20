using System.Text.Json;
using System.Text.Json.Serialization;
using SharePointOnlineManager.Data;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Service for managing and executing tasks.
/// </summary>
public class TaskService : ITaskService
{
    private readonly IDataStore<TaskDefinition> _taskStore;
    private readonly string _resultsFolder;
    private readonly JsonSerializerOptions _jsonOptions;

    public TaskService()
    {
        _taskStore = new JsonDataStore<TaskDefinition>(
            "tasks.json",
            t => t.Id,
            (t, id) => t.Id = id);

        _resultsFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SharePointOnlineManager",
            "results");

        Directory.CreateDirectory(_resultsFolder);

        _jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Converters = { new JsonStringEnumConverter() }
        };
    }

    public async Task<List<TaskDefinition>> GetAllTasksAsync()
    {
        var tasks = await _taskStore.GetAllAsync();
        return tasks.OrderByDescending(t => t.LastRunAt ?? t.CreatedAt).ToList();
    }

    public async Task<TaskDefinition?> GetTaskAsync(Guid id)
    {
        return await _taskStore.GetByIdAsync(id);
    }

    public async Task SaveTaskAsync(TaskDefinition task)
    {
        await _taskStore.SaveAsync(task);
    }

    public async Task DeleteTaskAsync(Guid id)
    {
        // Delete task results
        var resultFiles = Directory.GetFiles(_resultsFolder, $"{id}_*.json");
        foreach (var file in resultFiles)
        {
            try
            {
                File.Delete(file);
            }
            catch
            {
                // Ignore deletion errors
            }
        }

        await _taskStore.DeleteAsync(id);
    }

    public async Task<TaskResult> ExecuteTaskAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var result = new TaskResult
        {
            TaskId = task.Id,
            ExecutedAt = DateTime.UtcNow
        };

        task.Status = Models.TaskStatus.Running;
        task.LastRunAt = DateTime.UtcNow;
        task.LastError = null;
        await SaveTaskAsync(task);

        result.Log($"Starting task execution for {task.TargetSiteUrls.Count} sites");

        try
        {
            // Group sites by domain for efficient cookie usage
            var sitesByDomain = task.TargetSiteUrls
                .GroupBy(url => new Uri(url).Host)
                .ToList();

            int processedCount = 0;

            foreach (var domainGroup in sitesByDomain)
            {
                var domain = domainGroup.Key;
                result.Log($"Processing domain: {domain}");

                var cookies = authService.GetStoredCookies(domain);
                if (cookies == null || !cookies.IsValid)
                {
                    result.Log($"No valid credentials for {domain} - skipping {domainGroup.Count()} sites");

                    foreach (var siteUrl in domainGroup)
                    {
                        result.SiteResults.Add(new SiteListResult
                        {
                            SiteUrl = siteUrl,
                            Success = false,
                            ErrorMessage = "Authentication required"
                        });
                        result.FailedSites++;
                    }
                    processedCount += domainGroup.Count();
                    continue;
                }

                using var spService = new SharePointService(cookies);

                foreach (var siteUrl in domainGroup)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    processedCount++;
                    progress?.Report(new TaskProgress
                    {
                        CurrentSite = processedCount,
                        TotalSites = task.TargetSiteUrls.Count,
                        CurrentSiteUrl = siteUrl,
                        Message = $"Processing {processedCount}/{task.TargetSiteUrls.Count}: {siteUrl}"
                    });

                    result.Log($"Processing site: {siteUrl}");

                    var siteResult = new SiteListResult { SiteUrl = siteUrl };

                    try
                    {
                        // Get site info for title
                        var siteInfo = await spService.GetSiteInfoAsync(siteUrl);
                        siteResult.SiteTitle = siteInfo.Title;

                        // Get lists
                        var listsResult = await spService.GetListsAsync(siteUrl);

                        if (listsResult.IsSuccess && listsResult.Data != null)
                        {
                            siteResult.Lists = listsResult.Data;
                            siteResult.Success = true;
                            result.SuccessfulSites++;
                            result.Log($"  Found {listsResult.Data.Count} lists");
                        }
                        else if (listsResult.NeedsReauth)
                        {
                            siteResult.Success = false;
                            siteResult.ErrorMessage = "Authentication expired";
                            result.FailedSites++;
                            result.Log($"  Authentication expired");
                        }
                        else
                        {
                            siteResult.Success = false;
                            siteResult.ErrorMessage = listsResult.ErrorMessage ?? "Unknown error";
                            result.FailedSites++;
                            result.Log($"  Error: {siteResult.ErrorMessage}");
                        }
                    }
                    catch (Exception ex)
                    {
                        siteResult.Success = false;
                        siteResult.ErrorMessage = ex.Message;
                        result.FailedSites++;
                        result.Log($"  Exception: {ex.Message}");
                    }

                    result.SiteResults.Add(siteResult);
                    result.TotalSitesProcessed++;
                }
            }

            result.CompletedAt = DateTime.UtcNow;
            result.Success = result.FailedSites == 0;
            result.Log($"Task completed. Successful: {result.SuccessfulSites}, Failed: {result.FailedSites}");

            task.Status = result.FailedSites == 0 ? Models.TaskStatus.Completed : Models.TaskStatus.Failed;
            task.CompletedAt = DateTime.UtcNow;

            if (result.FailedSites > 0)
            {
                task.LastError = $"{result.FailedSites} site(s) failed";
            }
        }
        catch (OperationCanceledException)
        {
            result.CompletedAt = DateTime.UtcNow;
            result.Success = false;
            result.ErrorMessage = "Task was cancelled";
            result.Log("Task cancelled by user");

            task.Status = Models.TaskStatus.Cancelled;
            task.LastError = "Cancelled";
        }
        catch (Exception ex)
        {
            result.CompletedAt = DateTime.UtcNow;
            result.Success = false;
            result.ErrorMessage = ex.Message;
            result.Log($"Task failed: {ex.Message}");

            task.Status = Models.TaskStatus.Failed;
            task.LastError = ex.Message;
        }

        await SaveTaskAsync(task);
        await SaveTaskResultAsync(result);

        return result;
    }

    public async Task<List<TaskResult>> GetTaskResultsAsync(Guid taskId)
    {
        var results = new List<TaskResult>();
        var pattern = $"{taskId}_*.json";
        var files = Directory.GetFiles(_resultsFolder, pattern)
            .OrderByDescending(f => f);

        foreach (var file in files)
        {
            try
            {
                var json = await File.ReadAllTextAsync(file);
                var result = JsonSerializer.Deserialize<TaskResult>(json, _jsonOptions);
                if (result != null)
                {
                    results.Add(result);
                }
            }
            catch
            {
                // Skip invalid files
            }
        }

        return results;
    }

    public async Task<TaskResult?> GetLatestTaskResultAsync(Guid taskId)
    {
        var results = await GetTaskResultsAsync(taskId);
        return results.FirstOrDefault();
    }

    public async Task SaveTaskResultAsync(TaskResult result)
    {
        var timestamp = result.ExecutedAt.ToString("yyyyMMdd_HHmmss");
        var fileName = $"{result.TaskId}_{timestamp}.json";
        var filePath = Path.Combine(_resultsFolder, fileName);

        var json = JsonSerializer.Serialize(result, _jsonOptions);
        await File.WriteAllTextAsync(filePath, json);
    }

    public async Task<ListCompareResult> ExecuteListCompareAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IConnectionManager connectionManager,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var result = new ListCompareResult
        {
            TaskId = task.Id,
            ExecutedAt = DateTime.UtcNow
        };

        task.Status = Models.TaskStatus.Running;
        task.LastRunAt = DateTime.UtcNow;
        task.LastError = null;
        await SaveTaskAsync(task);

        // Deserialize configuration
        if (string.IsNullOrEmpty(task.ConfigurationJson))
        {
            result.Success = false;
            result.ErrorMessage = "Task configuration is missing";
            result.Log("Error: Task configuration is missing");
            task.Status = Models.TaskStatus.Failed;
            task.LastError = result.ErrorMessage;
            await SaveTaskAsync(task);
            await SaveListCompareResultAsync(result);
            return result;
        }

        ListCompareConfiguration config;
        try
        {
            config = JsonSerializer.Deserialize<ListCompareConfiguration>(task.ConfigurationJson, _jsonOptions)
                     ?? throw new InvalidOperationException("Failed to deserialize configuration");
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Invalid configuration: {ex.Message}";
            result.Log($"Error: {result.ErrorMessage}");
            task.Status = Models.TaskStatus.Failed;
            task.LastError = result.ErrorMessage;
            await SaveTaskAsync(task);
            await SaveListCompareResultAsync(result);
            return result;
        }

        result.Log($"Starting list compare task for {config.SitePairs.Count} site pairs");

        try
        {
            // Get connections
            var sourceConnection = await connectionManager.GetConnectionAsync(config.SourceConnectionId);
            var targetConnection = await connectionManager.GetConnectionAsync(config.TargetConnectionId);

            if (sourceConnection == null || targetConnection == null)
            {
                throw new InvalidOperationException("Source or target connection not found");
            }

            result.Log($"Source connection: {sourceConnection.Name}");
            result.Log($"Target connection: {targetConnection.Name}");

            // Build exclusion list
            var excludedLists = new HashSet<string>(config.ExcludedLists, StringComparer.OrdinalIgnoreCase);
            if (!config.IncludeSiteAssets)
            {
                foreach (var optExcluded in ListCompareConfiguration.OptionalExcludedLists)
                {
                    excludedLists.Add(optExcluded);
                }
            }

            result.Log($"Excluded lists: {excludedLists.Count}");

            // Get cookies for source connection - try tenant domain first, then admin domain
            var sourceCookies = authService.GetStoredCookies(sourceConnection.TenantDomain)
                                ?? authService.GetStoredCookies(sourceConnection.AdminDomain);
            if (sourceCookies == null || !sourceCookies.IsValid)
            {
                throw new InvalidOperationException(
                    $"No valid credentials for source tenant {sourceConnection.TenantName}. Please authenticate first.");
            }

            // Get cookies for target connection - try tenant domain first, then admin domain
            var targetCookies = authService.GetStoredCookies(targetConnection.TenantDomain)
                                ?? authService.GetStoredCookies(targetConnection.AdminDomain);
            if (targetCookies == null || !targetCookies.IsValid)
            {
                throw new InvalidOperationException(
                    $"No valid credentials for target tenant {targetConnection.TenantName}. Please authenticate first.");
            }

            // Create services once for the entire task (cookies work across tenant)
            // Use tenant domain (not admin domain) for the actual site requests
            using var sourceService = new SharePointService(sourceCookies, sourceConnection.TenantDomain);
            using var targetService = new SharePointService(targetCookies, targetConnection.TenantDomain);

            int processedCount = 0;

            foreach (var pair in config.SitePairs)
            {
                cancellationToken.ThrowIfCancellationRequested();

                processedCount++;
                progress?.Report(new TaskProgress
                {
                    CurrentSite = processedCount,
                    TotalSites = config.SitePairs.Count,
                    CurrentSiteUrl = pair.SourceUrl,
                    Message = $"Comparing {processedCount}/{config.SitePairs.Count}: {pair.SourceUrl}"
                });

                result.Log($"Processing pair {processedCount}: {pair.SourceUrl} <-> {pair.TargetUrl}");

                var siteResult = new SiteCompareResult
                {
                    SourceSiteUrl = pair.SourceUrl,
                    TargetSiteUrl = pair.TargetUrl
                };

                try
                {
                    // Get source site info and lists
                    var sourceInfo = await sourceService.GetSiteInfoAsync(pair.SourceUrl);
                    siteResult.SourceSiteTitle = sourceInfo.Title;

                    var sourceListsResult = await sourceService.GetListsAsync(pair.SourceUrl);
                    if (!sourceListsResult.IsSuccess || sourceListsResult.Data == null)
                    {
                        throw new InvalidOperationException(
                            $"Failed to get source lists: {sourceListsResult.ErrorMessage}");
                    }

                    // Get target site info and lists
                    var targetInfo = await targetService.GetSiteInfoAsync(pair.TargetUrl);
                    siteResult.TargetSiteTitle = targetInfo.Title;

                    var targetListsResult = await targetService.GetListsAsync(pair.TargetUrl);
                    if (!targetListsResult.IsSuccess || targetListsResult.Data == null)
                    {
                        throw new InvalidOperationException(
                            $"Failed to get target lists: {targetListsResult.ErrorMessage}");
                    }

                    // Filter lists
                    var sourceLists = FilterLists(sourceListsResult.Data, excludedLists, config.IncludeHiddenLists);
                    var targetLists = FilterLists(targetListsResult.Data, excludedLists, config.IncludeHiddenLists);

                    result.Log($"  Source: {sourceLists.Count} lists, Target: {targetLists.Count} lists");

                    // Compare lists by title (case-insensitive)
                    var sourceByTitle = sourceLists.ToDictionary(l => l.Title, StringComparer.OrdinalIgnoreCase);
                    var targetByTitle = targetLists.ToDictionary(l => l.Title, StringComparer.OrdinalIgnoreCase);

                    // Process all source lists
                    foreach (var sourceList in sourceLists)
                    {
                        var comparison = new ListCompareItem
                        {
                            SourceSiteUrl = pair.SourceUrl,
                            TargetSiteUrl = pair.TargetUrl,
                            SourceSiteTitle = siteResult.SourceSiteTitle,
                            TargetSiteTitle = siteResult.TargetSiteTitle,
                            ListTitle = sourceList.Title,
                            ListType = sourceList.ListType,
                            SourceListUrl = sourceList.GetAbsoluteUrl(pair.SourceUrl),
                            SourceCount = sourceList.ItemCount
                        };

                        if (targetByTitle.TryGetValue(sourceList.Title, out var targetList))
                        {
                            comparison.TargetCount = targetList.ItemCount;
                            comparison.TargetListUrl = targetList.GetAbsoluteUrl(pair.TargetUrl);
                            comparison.Status = IsWithinThreshold(
                                sourceList.ItemCount,
                                targetList.ItemCount,
                                config.ThresholdType,
                                config.ThresholdValue)
                                ? ListCompareStatus.Match
                                : ListCompareStatus.Mismatch;
                        }
                        else
                        {
                            comparison.TargetCount = 0;
                            comparison.Status = ListCompareStatus.SourceOnly;
                        }

                        siteResult.ListComparisons.Add(comparison);
                    }

                    // Process target-only lists
                    foreach (var targetList in targetLists)
                    {
                        if (!sourceByTitle.ContainsKey(targetList.Title))
                        {
                            siteResult.ListComparisons.Add(new ListCompareItem
                            {
                                SourceSiteUrl = pair.SourceUrl,
                                TargetSiteUrl = pair.TargetUrl,
                                SourceSiteTitle = siteResult.SourceSiteTitle,
                                TargetSiteTitle = siteResult.TargetSiteTitle,
                                ListTitle = targetList.Title,
                                ListType = targetList.ListType,
                                TargetListUrl = targetList.GetAbsoluteUrl(pair.TargetUrl),
                                SourceCount = 0,
                                TargetCount = targetList.ItemCount,
                                Status = ListCompareStatus.TargetOnly
                            });
                        }
                    }

                    siteResult.Success = true;
                    result.SuccessfulPairs++;

                    result.Log($"  Matches: {siteResult.MatchCount}, Mismatches: {siteResult.MismatchCount}, " +
                               $"Source Only: {siteResult.SourceOnlyCount}, Target Only: {siteResult.TargetOnlyCount}");
                }
                catch (Exception ex)
                {
                    siteResult.Success = false;
                    siteResult.ErrorMessage = ex.Message;
                    result.FailedPairs++;
                    result.Log($"  Error: {ex.Message}");
                }

                result.SiteResults.Add(siteResult);
                result.TotalPairsProcessed++;
            }

            result.CompletedAt = DateTime.UtcNow;
            result.Success = result.FailedPairs == 0;
            result.Log($"Task completed. Successful: {result.SuccessfulPairs}, Failed: {result.FailedPairs}");

            task.Status = result.FailedPairs == 0 ? Models.TaskStatus.Completed : Models.TaskStatus.Failed;
            task.CompletedAt = DateTime.UtcNow;

            if (result.FailedPairs > 0)
            {
                task.LastError = $"{result.FailedPairs} site pair(s) failed";
            }
        }
        catch (OperationCanceledException)
        {
            result.CompletedAt = DateTime.UtcNow;
            result.Success = false;
            result.ErrorMessage = "Task was cancelled";
            result.Log("Task cancelled by user");

            task.Status = Models.TaskStatus.Cancelled;
            task.LastError = "Cancelled";
        }
        catch (Exception ex)
        {
            result.CompletedAt = DateTime.UtcNow;
            result.Success = false;
            result.ErrorMessage = ex.Message;
            result.Log($"Task failed: {ex.Message}");

            task.Status = Models.TaskStatus.Failed;
            task.LastError = ex.Message;
        }

        await SaveTaskAsync(task);
        await SaveListCompareResultAsync(result);

        return result;
    }

    private static List<ListInfo> FilterLists(
        List<ListInfo> lists,
        HashSet<string> excludedLists,
        bool includeHidden)
    {
        return lists.Where(l =>
        {
            // Check if excluded
            if (excludedLists.Contains(l.Title))
                return false;

            // Check hidden
            if (!includeHidden && l.Hidden)
                return false;

            // Never exclude Site Pages
            if (ListCompareConfiguration.NeverExcludedLists.Contains(l.Title, StringComparer.OrdinalIgnoreCase))
                return true;

            return true;
        }).ToList();
    }

    private static bool IsWithinThreshold(
        int sourceCount,
        int targetCount,
        ThresholdType thresholdType,
        int thresholdValue)
    {
        if (sourceCount == targetCount)
            return true;

        if (thresholdType == ThresholdType.AbsoluteCount)
        {
            return Math.Abs(targetCount - sourceCount) <= thresholdValue;
        }
        else // Percentage
        {
            if (sourceCount == 0)
                return targetCount == 0;

            var percentDiff = Math.Abs((double)(targetCount - sourceCount) / sourceCount * 100);
            return percentDiff <= thresholdValue;
        }
    }

    public async Task<List<ListCompareResult>> GetListCompareResultsAsync(Guid taskId)
    {
        var results = new List<ListCompareResult>();
        var pattern = $"listcompare_{taskId}_*.json";
        var files = Directory.GetFiles(_resultsFolder, pattern)
            .OrderByDescending(f => f);

        foreach (var file in files)
        {
            try
            {
                var json = await File.ReadAllTextAsync(file);
                var result = JsonSerializer.Deserialize<ListCompareResult>(json, _jsonOptions);
                if (result != null)
                {
                    results.Add(result);
                }
            }
            catch
            {
                // Skip invalid files
            }
        }

        return results;
    }

    public async Task<ListCompareResult?> GetLatestListCompareResultAsync(Guid taskId)
    {
        var results = await GetListCompareResultsAsync(taskId);
        return results.FirstOrDefault();
    }

    public async Task SaveListCompareResultAsync(ListCompareResult result)
    {
        var timestamp = result.ExecutedAt.ToString("yyyyMMdd_HHmmss");
        var fileName = $"listcompare_{result.TaskId}_{timestamp}.json";
        var filePath = Path.Combine(_resultsFolder, fileName);

        var json = JsonSerializer.Serialize(result, _jsonOptions);
        await File.WriteAllTextAsync(filePath, json);
    }
}

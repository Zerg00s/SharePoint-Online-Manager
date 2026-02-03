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

                // Pass the actual target domain so cookies are set correctly
                using var spService = new SharePointService(cookies, domain);

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

    public async Task<DocumentReportResult> ExecuteDocumentReportAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var result = new DocumentReportResult
        {
            TaskId = task.Id,
            ExecutedAt = DateTime.UtcNow
        };

        task.Status = Models.TaskStatus.Running;
        task.LastRunAt = DateTime.UtcNow;
        task.LastError = null;
        await SaveTaskAsync(task);

        // Deserialize configuration
        DocumentReportConfiguration config;
        if (string.IsNullOrEmpty(task.ConfigurationJson))
        {
            // Use defaults if no config
            config = new DocumentReportConfiguration
            {
                ConnectionId = task.ConnectionId,
                TargetSiteUrls = task.TargetSiteUrls
            };
        }
        else
        {
            try
            {
                config = JsonSerializer.Deserialize<DocumentReportConfiguration>(task.ConfigurationJson, _jsonOptions)
                         ?? new DocumentReportConfiguration
                         {
                             ConnectionId = task.ConnectionId,
                             TargetSiteUrls = task.TargetSiteUrls
                         };
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Invalid configuration: {ex.Message}";
                result.Log($"Error: {result.ErrorMessage}");
                task.Status = Models.TaskStatus.Failed;
                task.LastError = result.ErrorMessage;
                await SaveTaskAsync(task);
                await SaveDocumentReportResultAsync(result);
                return result;
            }
        }

        // Use task.TargetSiteUrls if config doesn't have any
        var targetUrls = config.TargetSiteUrls.Count > 0 ? config.TargetSiteUrls : task.TargetSiteUrls;
        result.Log($"Starting document report task for {targetUrls.Count} sites");

        // Debug: List all stored credential domains
        var storedDomains = authService.GetStoredDomains();
        System.Diagnostics.Debug.WriteLine($"[SPOManager] ExecuteDocumentReportAsync - Stored credential domains: [{string.Join(", ", storedDomains)}]");
        result.Log($"DEBUG: Stored credential domains: [{string.Join(", ", storedDomains)}]");
        result.Log($"DEBUG: Target URLs: [{string.Join(", ", targetUrls)}]");

        // Parse extension filter
        var extensionFilter = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(config.ExtensionFilter))
        {
            var extensions = config.ExtensionFilter
                .Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim().TrimStart('.').ToLowerInvariant());
            foreach (var ext in extensions)
            {
                extensionFilter.Add(ext);
            }
            result.Log($"Extension filter: {string.Join(", ", extensionFilter)}");
        }

        try
        {
            // Group sites by domain for efficient cookie usage
            var sitesByDomain = targetUrls
                .GroupBy(url => new Uri(url).Host)
                .ToList();

            int processedCount = 0;

            foreach (var domainGroup in sitesByDomain)
            {
                var domain = domainGroup.Key;
                result.Log($"Processing domain: {domain}");
                System.Diagnostics.Debug.WriteLine($"[SPOManager] Looking for cookies for domain: {domain}");

                var cookies = authService.GetStoredCookies(domain);
                System.Diagnostics.Debug.WriteLine($"[SPOManager] GetStoredCookies result - Found: {cookies != null}, IsValid: {cookies?.IsValid}, Domain: {cookies?.Domain}, User: {cookies?.UserEmail}");
                result.Log($"DEBUG: Cookies lookup for '{domain}' - Found: {cookies != null}, IsValid: {cookies?.IsValid}");

                if (cookies == null || !cookies.IsValid)
                {
                    result.Log($"No valid credentials for {domain} - skipping {domainGroup.Count()} sites");
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] FAILED: No valid credentials for {domain}");

                    foreach (var siteUrl in domainGroup)
                    {
                        result.SiteResults.Add(new SiteDocumentResult
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

                System.Diagnostics.Debug.WriteLine($"[SPOManager] SUCCESS: Using cookies from {cookies.Domain} for {domain}");
                // Pass the actual target domain so cookies are set correctly
                using var spService = new SharePointService(cookies, domain);

                foreach (var siteUrl in domainGroup)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    processedCount++;
                    progress?.Report(new TaskProgress
                    {
                        CurrentSite = processedCount,
                        TotalSites = targetUrls.Count,
                        CurrentSiteUrl = siteUrl,
                        Message = $"Processing {processedCount}/{targetUrls.Count}: {siteUrl}"
                    });

                    result.Log($"Processing site: {siteUrl}");

                    var siteResult = new SiteDocumentResult { SiteUrl = siteUrl };

                    try
                    {
                        // Get site info for title
                        var siteInfo = await spService.GetSiteInfoAsync(siteUrl);
                        siteResult.SiteTitle = siteInfo.Title;

                        // Get document libraries (BaseTemplate = 101)
                        var listsResult = await spService.GetListsAsync(siteUrl, config.IncludeHiddenLibraries);

                        if (!listsResult.IsSuccess || listsResult.Data == null)
                        {
                            throw new InvalidOperationException(
                                $"Failed to get lists: {listsResult.ErrorMessage}");
                        }

                        var documentLibraries = listsResult.Data
                            .Where(l => l.BaseTemplate == 101) // Document Library template
                            .ToList();

                        result.Log($"  Found {documentLibraries.Count} document libraries");

                        foreach (var library in documentLibraries)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            result.Log($"  Processing library: {library.Title}");

                            var filesResult = await spService.GetDocumentLibraryFilesAsync(
                                siteUrl,
                                library.Title,
                                config.IncludeSubfolders,
                                config.IncludeVersionCount);

                            if (filesResult.IsSuccess && filesResult.Data != null)
                            {
                                var files = filesResult.Data;

                                // Apply extension filter if specified
                                if (extensionFilter.Count > 0)
                                {
                                    files = files.Where(f => extensionFilter.Contains(f.Extension)).ToList();
                                }

                                // Set site title on all documents
                                foreach (var file in files)
                                {
                                    file.SiteTitle = siteResult.SiteTitle;
                                }

                                siteResult.Documents.AddRange(files);
                                siteResult.LibrariesProcessed++;
                                result.Log($"    Found {files.Count} documents");
                            }
                            else
                            {
                                result.Log($"    Error getting files: {filesResult.ErrorMessage}");
                            }
                        }

                        siteResult.Success = true;
                        result.SuccessfulSites++;
                        result.Log($"  Site completed: {siteResult.TotalDocuments} documents, {siteResult.LibrariesProcessed} libraries");
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

            var (totalDocs, totalSize, totalLibs) = result.GetSummary();
            result.Log($"Task completed. Sites: {result.SuccessfulSites} successful, {result.FailedSites} failed");
            result.Log($"Total: {totalDocs} documents, {FormatSize(totalSize)}, {totalLibs} libraries");

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
        await SaveDocumentReportResultAsync(result);

        return result;
    }

    private static string FormatSize(long bytes)
    {
        string[] sizes = ["B", "KB", "MB", "GB", "TB"];
        double len = bytes;
        int order = 0;
        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len = len / 1024;
        }
        return $"{len:0.##} {sizes[order]}";
    }

    public async Task<List<DocumentReportResult>> GetDocumentReportResultsAsync(Guid taskId)
    {
        var results = new List<DocumentReportResult>();
        var pattern = $"docreport_{taskId}_*.json";
        var files = Directory.GetFiles(_resultsFolder, pattern)
            .OrderByDescending(f => f);

        foreach (var file in files)
        {
            try
            {
                var json = await File.ReadAllTextAsync(file);
                var result = JsonSerializer.Deserialize<DocumentReportResult>(json, _jsonOptions);
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

    public async Task<DocumentReportResult?> GetLatestDocumentReportResultAsync(Guid taskId)
    {
        var results = await GetDocumentReportResultsAsync(taskId);
        return results.FirstOrDefault();
    }

    public async Task SaveDocumentReportResultAsync(DocumentReportResult result)
    {
        var timestamp = result.ExecutedAt.ToString("yyyyMMdd_HHmmss");
        var fileName = $"docreport_{result.TaskId}_{timestamp}.json";
        var filePath = Path.Combine(_resultsFolder, fileName);

        var json = JsonSerializer.Serialize(result, _jsonOptions);
        await File.WriteAllTextAsync(filePath, json);
    }

    public async Task<PermissionReportResult> ExecutePermissionReportAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        System.Diagnostics.Debug.WriteLine($"[SPOManager] ExecutePermissionReportAsync - START");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   TaskId: {task.Id}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   TaskName: {task.Name}");
        System.Diagnostics.Debug.WriteLine($"[SPOManager]   TargetSites: {task.TargetSiteUrls.Count}");

        var result = new PermissionReportResult
        {
            TaskId = task.Id,
            ExecutedAt = DateTime.UtcNow
        };

        task.Status = Models.TaskStatus.Running;
        task.LastRunAt = DateTime.UtcNow;
        task.LastError = null;
        await SaveTaskAsync(task);

        // Deserialize configuration
        PermissionReportConfiguration config;
        if (string.IsNullOrEmpty(task.ConfigurationJson))
        {
            config = new PermissionReportConfiguration
            {
                ConnectionId = task.ConnectionId,
                TargetSiteUrls = task.TargetSiteUrls
            };
        }
        else
        {
            try
            {
                config = JsonSerializer.Deserialize<PermissionReportConfiguration>(task.ConfigurationJson, _jsonOptions)
                         ?? new PermissionReportConfiguration
                         {
                             ConnectionId = task.ConnectionId,
                             TargetSiteUrls = task.TargetSiteUrls
                         };
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Invalid configuration: {ex.Message}";
                result.Log($"Error: {result.ErrorMessage}");
                task.Status = Models.TaskStatus.Failed;
                task.LastError = result.ErrorMessage;
                await SaveTaskAsync(task);
                await SavePermissionReportResultAsync(result);
                return result;
            }
        }

        var targetUrls = config.TargetSiteUrls.Count > 0 ? config.TargetSiteUrls : task.TargetSiteUrls;
        result.Log($"Starting permission report task for {targetUrls.Count} sites");
        result.Log($"Options: Sites={config.IncludeSitePermissions}, Lists={config.IncludeListPermissions}, " +
                   $"Folders={config.IncludeFolderPermissions}, Items={config.IncludeItemPermissions}, " +
                   $"Inherited={config.IncludeInheritedPermissions}, Hidden={config.IncludeHiddenLists}");

        try
        {
            // Group sites by domain for efficient cookie usage
            var sitesByDomain = targetUrls
                .GroupBy(url => new Uri(url).Host)
                .ToList();

            int processedCount = 0;

            foreach (var domainGroup in sitesByDomain)
            {
                var domain = domainGroup.Key;
                result.Log($"Processing domain: {domain}");
                System.Diagnostics.Debug.WriteLine($"[SPOManager] Processing domain: {domain} ({domainGroup.Count()} sites)");

                var cookies = authService.GetStoredCookies(domain);
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Cookies: {(cookies == null ? "NULL" : $"Domain={cookies.Domain}, Valid={cookies.IsValid}")}");

                if (cookies == null || !cookies.IsValid)
                {
                    result.Log($"No valid credentials for {domain} - skipping {domainGroup.Count()} sites");
                    System.Diagnostics.Debug.WriteLine($"[SPOManager]   SKIPPING - no valid credentials");

                    foreach (var siteUrl in domainGroup)
                    {
                        result.SiteResults.Add(new SitePermissionResult
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

                // Pass the actual target domain so cookies are set correctly
                System.Diagnostics.Debug.WriteLine($"[SPOManager]   Creating SharePointService for domain: {domain}");
                using var spService = new SharePointService(cookies, domain);

                foreach (var siteUrl in domainGroup)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    processedCount++;
                    progress?.Report(new TaskProgress
                    {
                        CurrentSite = processedCount,
                        TotalSites = targetUrls.Count,
                        CurrentSiteUrl = siteUrl,
                        Message = $"Processing {processedCount}/{targetUrls.Count}: {siteUrl}"
                    });

                    result.Log($"Processing site: {siteUrl}");
                    System.Diagnostics.Debug.WriteLine($"[SPOManager] ========== Processing site: {siteUrl} ==========");

                    var siteResult = new SitePermissionResult { SiteUrl = siteUrl };

                    try
                    {
                        // Get site info for title
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]   Getting site info...");
                        var siteInfo = await spService.GetSiteInfoAsync(siteUrl);
                        siteResult.SiteTitle = siteInfo.Title;
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]   Site title: {siteInfo.Title}, Connected: {siteInfo.IsConnected}");

                        if (!siteInfo.IsConnected)
                        {
                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Site not connected! Error: {siteInfo.ErrorMessage}");
                        }

                        // Determine site collection URL (for context)
                        var siteCollectionUrl = GetSiteCollectionUrl(siteUrl);
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]   Site collection URL: {siteCollectionUrl}");

                        // Get site/web permissions
                        if (config.IncludeSitePermissions)
                        {
                            result.Log($"  Getting site permissions...");
                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Getting web permissions (includeInherited={config.IncludeInheritedPermissions})...");

                            var webPermsResult = await spService.GetWebPermissionsAsync(
                                siteUrl, siteCollectionUrl, config.IncludeInheritedPermissions);

                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Web permissions result: Status={webPermsResult.Status}, Count={webPermsResult.Data?.Count ?? 0}");

                            if (webPermsResult.IsSuccess && webPermsResult.Data != null)
                            {
                                siteResult.Permissions.AddRange(webPermsResult.Data);
                                result.Log($"    Found {webPermsResult.Data.Count} site permission entries");
                            }
                            else
                            {
                                result.Log($"    Error getting site permissions: {webPermsResult.ErrorMessage}");
                                System.Diagnostics.Debug.WriteLine($"[SPOManager]   ERROR: {webPermsResult.ErrorMessage}");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Skipping web permissions (config.IncludeSitePermissions=false)");
                        }

                        // Get lists and libraries
                        if (config.IncludeListPermissions || config.IncludeFolderPermissions || config.IncludeItemPermissions)
                        {
                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Getting lists (includeHidden={config.IncludeHiddenLists})...");
                            var listsResult = await spService.GetListsAsync(siteUrl, config.IncludeHiddenLists);
                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Lists result: Status={listsResult.Status}, Count={listsResult.Data?.Count ?? 0}");

                            if (listsResult.IsSuccess && listsResult.Data != null)
                            {
                                result.Log($"  Processing {listsResult.Data.Count} lists/libraries...");

                                int listIndex = 0;
                                foreach (var list in listsResult.Data)
                                {
                                    listIndex++;
                                    cancellationToken.ThrowIfCancellationRequested();

                                    var isLibrary = list.BaseTemplate == 101;
                                    System.Diagnostics.Debug.WriteLine($"[SPOManager]   --- List {listIndex}/{listsResult.Data.Count}: '{list.Title}' (Template={list.BaseTemplate}, IsLibrary={isLibrary}) ---");

                                    // Get list permissions
                                    if (config.IncludeListPermissions)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[SPOManager]     Getting list permissions...");
                                        var listPermsResult = await spService.GetListPermissionsAsync(
                                            siteUrl, siteCollectionUrl, list.Title, isLibrary,
                                            config.IncludeInheritedPermissions);

                                        System.Diagnostics.Debug.WriteLine($"[SPOManager]     List permissions result: Status={listPermsResult.Status}, Count={listPermsResult.Data?.Count ?? 0}");

                                        if (listPermsResult.IsSuccess && listPermsResult.Data != null && listPermsResult.Data.Count > 0)
                                        {
                                            siteResult.Permissions.AddRange(listPermsResult.Data);
                                            result.Log($"    {list.Title}: {listPermsResult.Data.Count} permission entries");
                                        }
                                    }

                                    // Get item/folder permissions (only if unique permissions exist)
                                    if (config.IncludeFolderPermissions || config.IncludeItemPermissions)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[SPOManager]     Getting item permissions (folders={config.IncludeFolderPermissions}, items={config.IncludeItemPermissions})...");
                                        var itemPermsResult = await spService.GetItemPermissionsAsync(
                                            siteUrl, siteCollectionUrl, list.Title, isLibrary,
                                            config.IncludeFolderPermissions, config.IncludeItemPermissions,
                                            config.IncludeInheritedPermissions);

                                        System.Diagnostics.Debug.WriteLine($"[SPOManager]     Item permissions result: Status={itemPermsResult.Status}, Count={itemPermsResult.Data?.Count ?? 0}");

                                        if (itemPermsResult.IsSuccess && itemPermsResult.Data != null && itemPermsResult.Data.Count > 0)
                                        {
                                            siteResult.Permissions.AddRange(itemPermsResult.Data);
                                            result.Log($"    {list.Title}: {itemPermsResult.Data.Count} item/folder permission entries");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[SPOManager]   ERROR getting lists: {listsResult.ErrorMessage}");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Skipping lists (config.IncludeListPermissions=false and folders/items disabled)");
                        }

                        siteResult.Success = true;
                        result.SuccessfulSites++;
                        result.Log($"  Site completed: {siteResult.TotalPermissions} permission entries");
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]   SITE COMPLETE: {siteResult.TotalPermissions} permission entries");
                    }
                    catch (Exception ex)
                    {
                        siteResult.Success = false;
                        siteResult.ErrorMessage = ex.Message;
                        result.FailedSites++;
                        result.Log($"  Exception: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]   SITE EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"[SPOManager]   StackTrace: {ex.StackTrace}");
                    }

                    result.SiteResults.Add(siteResult);
                    result.TotalSitesProcessed++;
                }
            }

            result.CompletedAt = DateTime.UtcNow;
            result.Success = result.FailedSites == 0;

            var (totalPerms, uniqueObjects, uniquePrincipals) = result.GetSummary();
            result.Log($"Task completed. Sites: {result.SuccessfulSites} successful, {result.FailedSites} failed");
            result.Log($"Total: {totalPerms} permissions, {uniqueObjects} unique objects, {uniquePrincipals} principals");

            System.Diagnostics.Debug.WriteLine($"[SPOManager] ========================================");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] TASK COMPLETED");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Successful sites: {result.SuccessfulSites}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Failed sites: {result.FailedSites}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Total permissions: {totalPerms}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Unique objects: {uniqueObjects}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Unique principals: {uniquePrincipals}");
            System.Diagnostics.Debug.WriteLine($"[SPOManager] ========================================");

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
        await SavePermissionReportResultAsync(result);

        return result;
    }

    private static string GetSiteCollectionUrl(string siteUrl)
    {
        // Extract site collection URL from site URL
        // e.g., https://tenant.sharepoint.com/sites/hr/subsite -> https://tenant.sharepoint.com/sites/hr
        var uri = new Uri(siteUrl);
        var pathParts = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);

        if (pathParts.Length >= 2 && (pathParts[0].Equals("sites", StringComparison.OrdinalIgnoreCase) ||
                                       pathParts[0].Equals("teams", StringComparison.OrdinalIgnoreCase)))
        {
            return $"{uri.Scheme}://{uri.Host}/{pathParts[0]}/{pathParts[1]}";
        }

        // Root site collection
        return $"{uri.Scheme}://{uri.Host}";
    }

    public async Task<List<PermissionReportResult>> GetPermissionReportResultsAsync(Guid taskId)
    {
        var results = new List<PermissionReportResult>();
        var pattern = $"permreport_{taskId}_*.json";
        var files = Directory.GetFiles(_resultsFolder, pattern)
            .OrderByDescending(f => f);

        foreach (var file in files)
        {
            try
            {
                var json = await File.ReadAllTextAsync(file);
                var result = JsonSerializer.Deserialize<PermissionReportResult>(json, _jsonOptions);
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

    public async Task<PermissionReportResult?> GetLatestPermissionReportResultAsync(Guid taskId)
    {
        var results = await GetPermissionReportResultsAsync(taskId);
        return results.FirstOrDefault();
    }

    public async Task SavePermissionReportResultAsync(PermissionReportResult result)
    {
        var timestamp = result.ExecutedAt.ToString("yyyyMMdd_HHmmss");
        var fileName = $"permreport_{result.TaskId}_{timestamp}.json";
        var filePath = Path.Combine(_resultsFolder, fileName);

        var json = JsonSerializer.Serialize(result, _jsonOptions);
        await File.WriteAllTextAsync(filePath, json);
    }

    public async Task<NavigationSettingsResult> ExecuteNavigationSettingsSyncAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IConnectionManager connectionManager,
        bool applyMode = false,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var result = new NavigationSettingsResult
        {
            TaskId = task.Id,
            ExecutedAt = DateTime.UtcNow,
            ApplyMode = applyMode
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
            await SaveNavigationSettingsResultAsync(result);
            return result;
        }

        NavigationSettingsConfiguration config;
        try
        {
            config = JsonSerializer.Deserialize<NavigationSettingsConfiguration>(task.ConfigurationJson, _jsonOptions)
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
            await SaveNavigationSettingsResultAsync(result);
            return result;
        }

        result.Log($"Starting navigation settings {(applyMode ? "sync" : "compare")} for {config.SitePairs.Count} site pairs");

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

            // Get cookies for source connection
            var sourceCookies = authService.GetStoredCookies(sourceConnection.TenantDomain)
                                ?? authService.GetStoredCookies(sourceConnection.AdminDomain);
            if (sourceCookies == null || !sourceCookies.IsValid)
            {
                throw new InvalidOperationException(
                    $"No valid credentials for source tenant {sourceConnection.TenantName}. Please authenticate first.");
            }

            // Get cookies for target connection
            var targetCookies = authService.GetStoredCookies(targetConnection.TenantDomain)
                                ?? authService.GetStoredCookies(targetConnection.AdminDomain);
            if (targetCookies == null || !targetCookies.IsValid)
            {
                throw new InvalidOperationException(
                    $"No valid credentials for target tenant {targetConnection.TenantName}. Please authenticate first.");
            }

            // Create services
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
                    Message = $"{(applyMode ? "Applying" : "Comparing")} {processedCount}/{config.SitePairs.Count}: {pair.SourceUrl}"
                });

                result.Log($"Processing pair {processedCount}: {pair.SourceUrl} <-> {pair.TargetUrl}");

                var siteResult = new NavigationSettingsCompareItem
                {
                    SourceSiteUrl = pair.SourceUrl,
                    TargetSiteUrl = pair.TargetUrl
                };

                try
                {
                    // Get source site info and navigation settings
                    var sourceInfo = await sourceService.GetSiteInfoAsync(pair.SourceUrl);
                    siteResult.SourceSiteTitle = sourceInfo.Title;

                    var sourceSettingsResult = await sourceService.GetNavigationSettingsAsync(pair.SourceUrl);
                    if (!sourceSettingsResult.IsSuccess || sourceSettingsResult.Data == null)
                    {
                        throw new InvalidOperationException(
                            $"Failed to get source navigation settings: {sourceSettingsResult.ErrorMessage}");
                    }

                    siteResult.SourceHorizontalQuickLaunch = sourceSettingsResult.Data.HorizontalQuickLaunch;
                    siteResult.SourceMegaMenuEnabled = sourceSettingsResult.Data.MegaMenuEnabled;

                    // Get target site info and navigation settings
                    var targetInfo = await targetService.GetSiteInfoAsync(pair.TargetUrl);
                    siteResult.TargetSiteTitle = targetInfo.Title;

                    var targetSettingsResult = await targetService.GetNavigationSettingsAsync(pair.TargetUrl);
                    if (!targetSettingsResult.IsSuccess || targetSettingsResult.Data == null)
                    {
                        throw new InvalidOperationException(
                            $"Failed to get target navigation settings: {targetSettingsResult.ErrorMessage}");
                    }

                    siteResult.TargetHorizontalQuickLaunch = targetSettingsResult.Data.HorizontalQuickLaunch;
                    siteResult.TargetMegaMenuEnabled = targetSettingsResult.Data.MegaMenuEnabled;

                    result.Log($"  Source: HQL={siteResult.SourceHorizontalQuickLaunch}, MegaMenu={siteResult.SourceMegaMenuEnabled}");
                    result.Log($"  Target: HQL={siteResult.TargetHorizontalQuickLaunch}, MegaMenu={siteResult.TargetMegaMenuEnabled}");

                    // Determine status
                    if (siteResult.AllSettingsMatch)
                    {
                        siteResult.Status = NavigationSettingsStatus.Match;
                        result.MatchingPairs++;
                        result.Log($"  Status: Match");
                    }
                    else
                    {
                        if (applyMode)
                        {
                            // Apply source settings to target
                            result.Log($"  Applying settings to target...");
                            var applyResult = await targetService.SetNavigationSettingsAsync(
                                pair.TargetUrl,
                                new NavigationSettings
                                {
                                    HorizontalQuickLaunch = siteResult.SourceHorizontalQuickLaunch,
                                    MegaMenuEnabled = siteResult.SourceMegaMenuEnabled
                                });

                            if (applyResult.IsSuccess && applyResult.Data)
                            {
                                siteResult.Status = NavigationSettingsStatus.Applied;
                                siteResult.TargetHorizontalQuickLaunch = siteResult.SourceHorizontalQuickLaunch;
                                siteResult.TargetMegaMenuEnabled = siteResult.SourceMegaMenuEnabled;
                                result.AppliedPairs++;
                                result.Log($"  Status: Applied successfully");
                            }
                            else
                            {
                                siteResult.Status = NavigationSettingsStatus.Failed;
                                siteResult.ErrorMessage = applyResult.ErrorMessage ?? "Failed to apply settings";
                                result.FailedPairs++;
                                result.Log($"  Status: Failed - {siteResult.ErrorMessage}");
                            }
                        }
                        else
                        {
                            siteResult.Status = NavigationSettingsStatus.Mismatch;
                            result.MismatchedPairs++;
                            result.Log($"  Status: Mismatch");
                        }
                    }
                }
                catch (Exception ex)
                {
                    siteResult.Status = NavigationSettingsStatus.Error;
                    siteResult.ErrorMessage = ex.Message;
                    result.FailedPairs++;
                    result.Log($"  Error: {ex.Message}");
                }

                result.SiteResults.Add(siteResult);
                result.TotalPairsProcessed++;
            }

            result.CompletedAt = DateTime.UtcNow;
            result.Success = result.FailedPairs == 0;

            if (applyMode)
            {
                result.Log($"Sync completed. Applied: {result.AppliedPairs}, Failed: {result.FailedPairs}");
            }
            else
            {
                result.Log($"Compare completed. Matches: {result.MatchingPairs}, Mismatches: {result.MismatchedPairs}");
            }

            task.Status = result.FailedPairs == 0 ? Models.TaskStatus.Completed : Models.TaskStatus.Failed;
            task.CompletedAt = DateTime.UtcNow;

            if (result.FailedPairs > 0)
            {
                task.LastError = $"{result.FailedPairs} site(s) failed";
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
        await SaveNavigationSettingsResultAsync(result);

        return result;
    }

    public async Task<List<NavigationSettingsResult>> GetNavigationSettingsResultsAsync(Guid taskId)
    {
        var results = new List<NavigationSettingsResult>();
        var pattern = $"navsettings_{taskId}_*.json";
        var files = Directory.GetFiles(_resultsFolder, pattern)
            .OrderByDescending(f => f);

        foreach (var file in files)
        {
            try
            {
                var json = await File.ReadAllTextAsync(file);
                var result = JsonSerializer.Deserialize<NavigationSettingsResult>(json, _jsonOptions);
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

    public async Task<NavigationSettingsResult?> GetLatestNavigationSettingsResultAsync(Guid taskId)
    {
        var results = await GetNavigationSettingsResultsAsync(taskId);
        return results.FirstOrDefault();
    }

    public async Task SaveNavigationSettingsResultAsync(NavigationSettingsResult result)
    {
        var timestamp = result.ExecutedAt.ToString("yyyyMMdd_HHmmss");
        var fileName = $"navsettings_{result.TaskId}_{timestamp}.json";
        var filePath = Path.Combine(_resultsFolder, fileName);

        var json = JsonSerializer.Serialize(result, _jsonOptions);
        await File.WriteAllTextAsync(filePath, json);
    }
}

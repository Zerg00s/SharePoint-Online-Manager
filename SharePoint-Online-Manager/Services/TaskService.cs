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
    private readonly string _cacheFolder;
    private readonly JsonSerializerOptions _jsonOptions;

    public TaskService()
    {
        _taskStore = new JsonDataStore<TaskDefinition>(
            "tasks.json",
            t => t.Id,
            (t, id) => t.Id = id);

        var appDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SharePointOnlineManager");

        _resultsFolder = Path.Combine(appDataFolder, "results");
        _cacheFolder = Path.Combine(appDataFolder, "cache", "DocumentCompare");

        Directory.CreateDirectory(_resultsFolder);
        Directory.CreateDirectory(_cacheFolder);

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
                    var siteLabel = $"Site {processedCount}/{targetUrls.Count}";

                    progress?.Report(new TaskProgress
                    {
                        CurrentSite = processedCount,
                        TotalSites = targetUrls.Count,
                        CurrentSiteUrl = siteUrl,
                        Message = $"{siteLabel}: Connecting to {siteUrl}..."
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
                            progress?.Report(new TaskProgress
                            {
                                CurrentSite = processedCount,
                                TotalSites = targetUrls.Count,
                                CurrentSiteUrl = siteUrl,
                                Message = $"{siteLabel}: Scanning site permissions... ({siteResult.TotalPermissions} found so far)"
                            });

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
                            progress?.Report(new TaskProgress
                            {
                                CurrentSite = processedCount,
                                TotalSites = targetUrls.Count,
                                CurrentSiteUrl = siteUrl,
                                Message = $"{siteLabel}: Fetching lists/libraries..."
                            });

                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Getting lists (includeHidden={config.IncludeHiddenLists})...");
                            var listsResult = await spService.GetListsAsync(siteUrl, config.IncludeHiddenLists);
                            System.Diagnostics.Debug.WriteLine($"[SPOManager]   Lists result: Status={listsResult.Status}, Count={listsResult.Data?.Count ?? 0}");

                            if (listsResult.IsSuccess && listsResult.Data != null)
                            {
                                var totalLists = listsResult.Data.Count;
                                result.Log($"  Found {totalLists} lists/libraries");

                                progress?.Report(new TaskProgress
                                {
                                    CurrentSite = processedCount,
                                    TotalSites = targetUrls.Count,
                                    CurrentSiteUrl = siteUrl,
                                    Message = $"{siteLabel}: Found {totalLists} lists/libraries"
                                });

                                int listIndex = 0;
                                foreach (var list in listsResult.Data)
                                {
                                    listIndex++;
                                    cancellationToken.ThrowIfCancellationRequested();

                                    var isLibrary = list.BaseTemplate == 101;
                                    var listLabel = $"List {listIndex}/{totalLists} - {list.Title}";
                                    System.Diagnostics.Debug.WriteLine($"[SPOManager]   --- List {listIndex}/{totalLists}: '{list.Title}' (Template={list.BaseTemplate}, IsLibrary={isLibrary}) ---");

                                    // Get list permissions
                                    if (config.IncludeListPermissions)
                                    {
                                        progress?.Report(new TaskProgress
                                        {
                                            CurrentSite = processedCount,
                                            TotalSites = targetUrls.Count,
                                            CurrentSiteUrl = siteUrl,
                                            Message = $"{siteLabel}: {listLabel} (permissions)... [{siteResult.TotalPermissions} found]"
                                        });

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
                                        progress?.Report(new TaskProgress
                                        {
                                            CurrentSite = processedCount,
                                            TotalSites = targetUrls.Count,
                                            CurrentSiteUrl = siteUrl,
                                            Message = $"{siteLabel}: {listLabel} (scanning items)... [{siteResult.TotalPermissions} found]"
                                        });

                                        System.Diagnostics.Debug.WriteLine($"[SPOManager]     Getting item permissions (folders={config.IncludeFolderPermissions}, items={config.IncludeItemPermissions})...");
                                        var itemPermsResult = await spService.GetItemPermissionsAsync(
                                            siteUrl, siteCollectionUrl, list.Title, isLibrary,
                                            config.IncludeFolderPermissions, config.IncludeItemPermissions,
                                            config.IncludeInheritedPermissions,
                                            onPageScanned: (itemsScanned, uniquePerms) =>
                                            {
                                                progress?.Report(new TaskProgress
                                                {
                                                    CurrentSite = processedCount,
                                                    TotalSites = targetUrls.Count,
                                                    CurrentSiteUrl = siteUrl,
                                                    Message = $"{siteLabel}: {listLabel} ({itemsScanned} items scanned, {uniquePerms} unique) [{siteResult.TotalPermissions} found]"
                                                });
                                            });

                                        System.Diagnostics.Debug.WriteLine($"[SPOManager]     Item permissions result: Status={itemPermsResult.Status}, Count={itemPermsResult.Data?.Count ?? 0}");

                                        if (itemPermsResult.IsSuccess && itemPermsResult.Data != null)
                                        {
                                            if (itemPermsResult.Data.Count > 0)
                                            {
                                                siteResult.Permissions.AddRange(itemPermsResult.Data);
                                                var folders = itemPermsResult.Data.Count(p => p.ObjectType == PermissionObjectType.Folder);
                                                var docs = itemPermsResult.Data.Count(p => p.ObjectType == PermissionObjectType.Document);
                                                var listItems = itemPermsResult.Data.Count(p => p.ObjectType == PermissionObjectType.ListItem);
                                                result.Log($"    {list.Title}: {itemPermsResult.Data.Count} item/folder entries (Folders: {folders}, Documents: {docs}, ListItems: {listItems})");
                                            }
                                            else
                                            {
                                                result.Log($"    {list.Title}: 0 items with unique permissions");
                                            }
                                        }
                                        else if (!itemPermsResult.IsSuccess)
                                        {
                                            result.Log($"    {list.Title}: item scan error - {itemPermsResult.ErrorMessage}");
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
        CancellationToken cancellationToken = default,
        Action<NavigationSettingsCompareItem>? onSiteCompleted = null)
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
                onSiteCompleted?.Invoke(siteResult);
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

    public async Task<DocumentCompareResult> ExecuteDocumentCompareAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IConnectionManager connectionManager,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default,
        bool continueFromPrevious = false,
        Func<string, string, Task<AuthCookies?>>? reauthCallback = null)
    {
        var result = new DocumentCompareResult
        {
            TaskId = task.Id,
            ExecutedAt = DateTime.UtcNow
        };

        // Load previous result if continuing
        DocumentCompareResult? previousResult = null;
        HashSet<string> completedPairKeys = new(StringComparer.OrdinalIgnoreCase);

        if (continueFromPrevious)
        {
            previousResult = await GetLatestDocumentCompareResultAsync(task.Id);
            if (previousResult != null)
            {
                // Build set of completed source URLs (we use source URL as the unique key)
                foreach (var siteResult in previousResult.SiteResults.Where(s => s.Success))
                {
                    completedPairKeys.Add(siteResult.SourceSiteUrl);
                }

                // Copy previous results to new result
                result.SiteResults.AddRange(previousResult.SiteResults.Where(s => s.Success));
                result.SuccessfulPairs = previousResult.SuccessfulPairs;
                result.TotalPairsProcessed = previousResult.SiteResults.Count(s => s.Success);
                result.ThrottleRetryCount = previousResult.ThrottleRetryCount;

                // Copy execution log
                foreach (var logEntry in previousResult.ExecutionLog)
                {
                    result.ExecutionLog.Add(logEntry);
                }
                result.Log($"--- Continuing from previous run ({completedPairKeys.Count} pairs already completed) ---");
            }
        }

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
            await SaveDocumentCompareResultAsync(result);
            return result;
        }

        DocumentCompareConfiguration config;
        try
        {
            config = JsonSerializer.Deserialize<DocumentCompareConfiguration>(task.ConfigurationJson, _jsonOptions)
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
            await SaveDocumentCompareResultAsync(result);
            return result;
        }

        result.Log($"Starting document compare task for {config.SitePairs.Count} site pairs");

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
            var excludedLibraries = new HashSet<string>(config.ExcludedLibraries, StringComparer.OrdinalIgnoreCase);
            foreach (var defaultExcluded in DocumentCompareConfiguration.DefaultExcludedLibraries)
            {
                excludedLibraries.Add(defaultExcluded);
            }

            result.Log($"Excluded libraries: {excludedLibraries.Count}");

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

            // Create services (manually disposable for re-auth support)
            var sourceService = new SharePointService(sourceCookies, sourceConnection.TenantDomain);
            var targetService = new SharePointService(targetCookies, targetConnection.TenantDomain);

            // Track whether re-auth was declined per tenant
            bool sourceReauthDeclined = false;
            bool targetReauthDeclined = false;

            int processedCount = 0;

            try
            {

            foreach (var pair in config.SitePairs)
            {
                cancellationToken.ThrowIfCancellationRequested();

                processedCount++;

                // Skip if already completed in previous run
                if (completedPairKeys.Contains(pair.SourceUrl))
                {
                    progress?.Report(new TaskProgress
                    {
                        CurrentSite = processedCount,
                        TotalSites = config.SitePairs.Count,
                        CurrentSiteUrl = pair.SourceUrl,
                        Message = $"Skipping {processedCount}/{config.SitePairs.Count} (already completed): {pair.SourceUrl}"
                    });
                    continue;
                }

                progress?.Report(new TaskProgress
                {
                    CurrentSite = processedCount,
                    TotalSites = config.SitePairs.Count,
                    CurrentSiteUrl = pair.SourceUrl,
                    Message = $"Comparing {processedCount}/{config.SitePairs.Count}: {pair.SourceUrl}"
                });

                result.Log($"Processing pair {processedCount}: {pair.SourceUrl} <-> {pair.TargetUrl}");

                var siteResult = new SiteDocumentCompareResult
                {
                    SourceSiteUrl = pair.SourceUrl,
                    TargetSiteUrl = pair.TargetUrl
                };

                try
                {
                    // Get source site info
                    var sourceInfo = await sourceService.GetSiteInfoAsync(pair.SourceUrl);

                    // Check for source auth failure
                    if (!sourceInfo.IsConnected && IsAuthError(sourceInfo.ErrorMessage))
                    {
                        if (!sourceReauthDeclined && reauthCallback != null)
                        {
                            result.Log($"  Authentication failed for {sourceConnection.TenantDomain} — prompting re-authentication...");
                            var freshCookies = await reauthCallback(sourceConnection.TenantName, sourceConnection.TenantDomain);
                            if (freshCookies != null)
                            {
                                sourceService.Dispose();
                                sourceService = new SharePointService(freshCookies, sourceConnection.TenantDomain);
                                result.Log($"  Re-authenticated to {sourceConnection.TenantDomain}, retrying...");
                                sourceInfo = await sourceService.GetSiteInfoAsync(pair.SourceUrl);
                            }
                            else
                            {
                                sourceReauthDeclined = true;
                                result.Log($"  Re-authentication declined for {sourceConnection.TenantDomain}");
                            }
                        }
                    }

                    siteResult.SourceSiteTitle = sourceInfo.Title;

                    // Get target site info
                    var targetInfo = await targetService.GetSiteInfoAsync(pair.TargetUrl);

                    // Check for target auth failure
                    if (!targetInfo.IsConnected && IsAuthError(targetInfo.ErrorMessage))
                    {
                        if (!targetReauthDeclined && reauthCallback != null)
                        {
                            result.Log($"  Authentication failed for {targetConnection.TenantDomain} — prompting re-authentication...");
                            var freshCookies = await reauthCallback(targetConnection.TenantName, targetConnection.TenantDomain);
                            if (freshCookies != null)
                            {
                                targetService.Dispose();
                                targetService = new SharePointService(freshCookies, targetConnection.TenantDomain);
                                result.Log($"  Re-authenticated to {targetConnection.TenantDomain}, retrying...");
                                targetInfo = await targetService.GetSiteInfoAsync(pair.TargetUrl);
                            }
                            else
                            {
                                targetReauthDeclined = true;
                                result.Log($"  Re-authentication declined for {targetConnection.TenantDomain}");
                            }
                        }
                    }

                    siteResult.TargetSiteTitle = targetInfo.Title;

                    // Get document libraries from source (BaseTemplate = 101)
                    var sourceListsResult = await sourceService.GetListsAsync(pair.SourceUrl, config.IncludeHiddenLibraries);

                    // Check for source auth failure on GetListsAsync
                    if (sourceListsResult.NeedsReauth || sourceListsResult.Status == SharePointResultStatus.AccessDenied)
                    {
                        if (!sourceReauthDeclined && reauthCallback != null)
                        {
                            result.Log($"  Authentication failed for {sourceConnection.TenantDomain} — prompting re-authentication...");
                            var freshCookies = await reauthCallback(sourceConnection.TenantName, sourceConnection.TenantDomain);
                            if (freshCookies != null)
                            {
                                sourceService.Dispose();
                                sourceService = new SharePointService(freshCookies, sourceConnection.TenantDomain);
                                result.Log($"  Re-authenticated to {sourceConnection.TenantDomain}, retrying...");
                                sourceListsResult = await sourceService.GetListsAsync(pair.SourceUrl, config.IncludeHiddenLibraries);
                            }
                            else
                            {
                                sourceReauthDeclined = true;
                                result.Log($"  Re-authentication declined for {sourceConnection.TenantDomain}");
                            }
                        }
                    }

                    if (!sourceListsResult.IsSuccess || sourceListsResult.Data == null)
                    {
                        throw new InvalidOperationException(
                            $"Failed to get source libraries: {sourceListsResult.ErrorMessage}");
                    }

                    var sourceLibraries = sourceListsResult.Data
                        .Where(l => l.BaseTemplate == 101 && !excludedLibraries.Contains(l.Title))
                        .ToList();

                    result.Log($"  Source: {sourceLibraries.Count} document libraries");

                    // Get document libraries from target
                    var targetListsResult = await targetService.GetListsAsync(pair.TargetUrl, config.IncludeHiddenLibraries);

                    // Check for target auth failure on GetListsAsync
                    if (targetListsResult.NeedsReauth || targetListsResult.Status == SharePointResultStatus.AccessDenied)
                    {
                        if (!targetReauthDeclined && reauthCallback != null)
                        {
                            result.Log($"  Authentication failed for {targetConnection.TenantDomain} — prompting re-authentication...");
                            var freshCookies = await reauthCallback(targetConnection.TenantName, targetConnection.TenantDomain);
                            if (freshCookies != null)
                            {
                                targetService.Dispose();
                                targetService = new SharePointService(freshCookies, targetConnection.TenantDomain);
                                result.Log($"  Re-authenticated to {targetConnection.TenantDomain}, retrying...");
                                targetListsResult = await targetService.GetListsAsync(pair.TargetUrl, config.IncludeHiddenLibraries);
                            }
                            else
                            {
                                targetReauthDeclined = true;
                                result.Log($"  Re-authentication declined for {targetConnection.TenantDomain}");
                            }
                        }
                    }

                    if (!targetListsResult.IsSuccess || targetListsResult.Data == null)
                    {
                        throw new InvalidOperationException(
                            $"Failed to get target libraries: {targetListsResult.ErrorMessage}");
                    }

                    var targetLibraries = targetListsResult.Data
                        .Where(l => l.BaseTemplate == 101 && !excludedLibraries.Contains(l.Title))
                        .ToDictionary(l => l.Title, StringComparer.OrdinalIgnoreCase);

                    // Process each source library
                    int libraryIndex = 0;
                    foreach (var sourceLibrary in sourceLibraries)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        libraryIndex++;

                        siteResult.LibrariesProcessed++;
                        var libraryItemCount = sourceLibrary.ItemCount;
                        var sizeInfo = libraryItemCount > 0 ? $" ({libraryItemCount:N0} items)" : "";
                        result.Log($"  Processing library: {sourceLibrary.Title}{sizeInfo}");

                        // Check if library exists on target first (quick lookup)
                        var libraryExistsOnTarget = targetLibraries.ContainsKey(sourceLibrary.Title);
                        var targetLibrary = libraryExistsOnTarget ? targetLibraries[sourceLibrary.Title] : null;
                        var targetSizeInfo = targetLibrary?.ItemCount > 0 ? $" ({targetLibrary.ItemCount:N0} items)" : "";

                        // Report library-level progress
                        progress?.Report(new TaskProgress
                        {
                            CurrentSite = processedCount,
                            TotalSites = config.SitePairs.Count,
                            CurrentSiteUrl = pair.SourceUrl,
                            Message = $"Site {processedCount}/{config.SitePairs.Count} - Library {libraryIndex}/{sourceLibraries.Count}: {sourceLibrary.Title}{sizeInfo}"
                        });

                        // Try to get documents from cache first if enabled
                        List<DocumentCompareSourceItem>? sourceDocs = null;
                        List<DocumentCompareSourceItem>? targetDocs = null;
                        bool sourceFromCache = false;
                        bool targetFromCache = false;

                        if (config.UseCache)
                        {
                            var sourceCacheEntry = TryGetCachedDocuments(pair.SourceUrl, sourceLibrary.Title, config.CacheExpirationHours);
                            if (sourceCacheEntry != null)
                            {
                                sourceDocs = sourceCacheEntry.Documents;
                                // Re-compute relative paths using current extraction logic
                                // (cached paths may be stale if algorithm changed since caching)
                                SharePointService.RecomputeRelativePaths(sourceDocs, sourceLibrary.Title, pair.SourceUrl);
                                sourceFromCache = true;
                            }

                            if (libraryExistsOnTarget)
                            {
                                var targetCacheEntry = TryGetCachedDocuments(pair.TargetUrl, sourceLibrary.Title, config.CacheExpirationHours);
                                if (targetCacheEntry != null)
                                {
                                    targetDocs = targetCacheEntry.Documents;
                                    SharePointService.RecomputeRelativePaths(targetDocs, sourceLibrary.Title, pair.TargetUrl);
                                    targetFromCache = true;
                                }
                            }
                        }

                        // Fetch any documents not in cache
                        if (sourceDocs == null || (libraryExistsOnTarget && targetDocs == null))
                        {
                            var cacheStatus = config.UseCache
                                ? $" (cache: src={sourceFromCache}, tgt={targetFromCache})"
                                : "";
                            result.Log($"    Fetching source{sizeInfo} and target{targetSizeInfo} in parallel...{cacheStatus}");

                            // Only fetch what we don't have cached
                            var sourceDocsTask = sourceDocs == null
                                ? sourceService.GetDocumentsForCompareAsync(
                                    pair.SourceUrl,
                                    sourceLibrary.Title,
                                    config.IncludeAspxPages,
                                    new Progress<string>(msg => result.Log($"    Source: {msg}")),
                                    cancellationToken)
                                : Task.FromResult(new SharePointResult<List<DocumentCompareSourceItem>>
                                {
                                    Data = sourceDocs,
                                    Status = SharePointResultStatus.Success
                                });

                            var targetDocsTask = !libraryExistsOnTarget
                                ? Task.FromResult(new SharePointResult<List<DocumentCompareSourceItem>>
                                {
                                    Data = null,
                                    Status = SharePointResultStatus.NotFound,
                                    ErrorMessage = "Library not found on target"
                                })
                                : targetDocs != null
                                    ? Task.FromResult(new SharePointResult<List<DocumentCompareSourceItem>>
                                    {
                                        Data = targetDocs,
                                        Status = SharePointResultStatus.Success
                                    })
                                    : targetService.GetDocumentsForCompareAsync(
                                        pair.TargetUrl,
                                        sourceLibrary.Title,
                                        config.IncludeAspxPages,
                                        new Progress<string>(msg => result.Log($"    Target: {msg}")),
                                        cancellationToken);

                            await Task.WhenAll(sourceDocsTask, targetDocsTask);

                            var sourceDocsResult = await sourceDocsTask;
                            var targetDocsResult = await targetDocsTask;

                            if (!sourceDocsResult.IsSuccess || sourceDocsResult.Data == null)
                            {
                                result.Log($"    Error getting source docs: {sourceDocsResult.ErrorMessage}");
                                continue;
                            }

                            sourceDocs = sourceDocsResult.Data;

                            // Save to cache if we fetched fresh data
                            if (config.UseCache && !sourceFromCache)
                            {
                                SaveToCache(pair.SourceUrl, sourceLibrary.Title, sourceDocs);
                            }

                            if (targetDocsResult.IsSuccess && targetDocsResult.Data != null)
                            {
                                targetDocs = targetDocsResult.Data;
                                if (config.UseCache && !targetFromCache)
                                {
                                    SaveToCache(pair.TargetUrl, sourceLibrary.Title, targetDocs);
                                }
                            }
                        }
                        else
                        {
                            result.Log($"    Using cached data for source and target");
                        }

                        var cacheIndicator = sourceFromCache ? " (cached)" : "";
                        result.Log($"    Source: {sourceDocs.Count} documents{cacheIndicator}");

                        // Check if library exists on target
                        if (!libraryExistsOnTarget)
                        {
                            // All documents are SourceOnly
                            foreach (var sourceDoc in sourceDocs)
                            {
                                siteResult.DocumentComparisons.Add(new DocumentCompareItem
                                {
                                    SourceSiteUrl = pair.SourceUrl,
                                    TargetSiteUrl = pair.TargetUrl,
                                    LibraryName = sourceLibrary.Title,
                                    ItemType = sourceDoc.ItemType,
                                    FileName = sourceDoc.FileName,
                                    FileExtension = sourceDoc.ItemType == DocumentCompareItemType.Folder ? "" : Path.GetExtension(sourceDoc.FileName).TrimStart('.'),
                                    RelativePath = sourceDoc.RelativePath,
                                    SourceItemId = sourceDoc.Id,
                                    SourceSizeBytes = sourceDoc.SizeBytes,
                                    SourceVersionCount = sourceDoc.VersionCount,
                                    SourceAbsolutePath = BuildAbsoluteUrl(pair.SourceUrl, sourceDoc.ServerRelativeUrl),
                                    SourceCreated = sourceDoc.Created,
                                    SourceModified = sourceDoc.Modified,
                                    Status = DocumentCompareStatus.SourceOnly
                                });
                            }
                            result.Log($"    Library not found on target - {sourceDocs.Count} source-only items");
                            continue;
                        }

                        if (targetDocs == null)
                        {
                            result.Log($"    Error getting target docs - treating as source-only");
                            // Treat as all source-only
                            foreach (var sourceDoc in sourceDocs)
                            {
                                siteResult.DocumentComparisons.Add(new DocumentCompareItem
                                {
                                    SourceSiteUrl = pair.SourceUrl,
                                    TargetSiteUrl = pair.TargetUrl,
                                    LibraryName = sourceLibrary.Title,
                                    ItemType = sourceDoc.ItemType,
                                    FileName = sourceDoc.FileName,
                                    FileExtension = sourceDoc.ItemType == DocumentCompareItemType.Folder ? "" : Path.GetExtension(sourceDoc.FileName).TrimStart('.'),
                                    RelativePath = sourceDoc.RelativePath,
                                    SourceItemId = sourceDoc.Id,
                                    SourceSizeBytes = sourceDoc.SizeBytes,
                                    SourceVersionCount = sourceDoc.VersionCount,
                                    SourceAbsolutePath = BuildAbsoluteUrl(pair.SourceUrl, sourceDoc.ServerRelativeUrl),
                                    SourceCreated = sourceDoc.Created,
                                    SourceModified = sourceDoc.Modified,
                                    Status = DocumentCompareStatus.SourceOnly
                                });
                            }
                            continue;
                        }

                        var targetCacheIndicator = targetFromCache ? " (cached)" : "";
                        result.Log($"    Target: {targetDocs.Count} documents{targetCacheIndicator}");

                        // Deduplicate source docs by relative path (SharePoint API may return duplicates)
                        var sourceDeduped = new Dictionary<string, DocumentCompareSourceItem>(StringComparer.OrdinalIgnoreCase);
                        foreach (var doc in sourceDocs)
                        {
                            sourceDeduped.TryAdd(doc.RelativePath, doc);
                        }
                        if (sourceDeduped.Count < sourceDocs.Count)
                        {
                            result.Log($"    Removed {sourceDocs.Count - sourceDeduped.Count} duplicate source entries");
                            sourceDocs = sourceDeduped.Values.ToList();
                        }

                        // Index target docs by relative path (case-insensitive), handling duplicates
                        var targetByPath = new Dictionary<string, DocumentCompareSourceItem>(StringComparer.OrdinalIgnoreCase);
                        foreach (var doc in targetDocs)
                        {
                            if (!targetByPath.TryAdd(doc.RelativePath, doc))
                            {
                                result.Log($"    WARNING: Duplicate target file path: {doc.RelativePath}");
                            }
                        }

                        // Build secondary index using normalized paths (ShareGate character replacement)
                        // ShareGate often replaces these characters with underscore: " * : < > ? / \ & # % { } ~
                        // Only build if normalization is enabled
                        Dictionary<string, DocumentCompareSourceItem>? targetByNormalizedPath = null;
                        if (config.UseShareGateNormalization)
                        {
                            targetByNormalizedPath = new Dictionary<string, DocumentCompareSourceItem>(StringComparer.OrdinalIgnoreCase);
                            foreach (var doc in targetDocs)
                            {
                                var normalizedPath = NormalizePathForShareGate(doc.RelativePath);
                                targetByNormalizedPath.TryAdd(normalizedPath, doc);
                            }
                        }

                        // Compare documents
                        int foundCount = 0, sizeIssueCount = 0, sourceOnlyCount = 0;

                        foreach (var sourceDoc in sourceDocs)
                        {
                            var comparison = new DocumentCompareItem
                            {
                                SourceSiteUrl = pair.SourceUrl,
                                TargetSiteUrl = pair.TargetUrl,
                                LibraryName = sourceLibrary.Title,
                                ItemType = sourceDoc.ItemType,
                                FileName = sourceDoc.FileName,
                                FileExtension = sourceDoc.ItemType == DocumentCompareItemType.Folder ? "" : Path.GetExtension(sourceDoc.FileName).TrimStart('.'),
                                RelativePath = sourceDoc.RelativePath,
                                SourceItemId = sourceDoc.Id,
                                SourceSizeBytes = sourceDoc.SizeBytes,
                                SourceVersionCount = sourceDoc.VersionCount,
                                SourceAbsolutePath = BuildAbsoluteUrl(pair.SourceUrl, sourceDoc.ServerRelativeUrl),
                                SourceCreated = sourceDoc.Created,
                                SourceModified = sourceDoc.Modified
                            };

                            // Try exact match first
                            DocumentCompareSourceItem? targetDoc = null;
                            string? matchedTargetPath = null;

                            if (targetByPath.TryGetValue(sourceDoc.RelativePath, out targetDoc))
                            {
                                matchedTargetPath = sourceDoc.RelativePath;
                            }
                            // Try normalized path match only if enabled in config
                            else if (config.UseShareGateNormalization && targetByNormalizedPath != null)
                            {
                                var normalizedSourcePath = NormalizePathForShareGate(sourceDoc.RelativePath);
                                if (targetByNormalizedPath.TryGetValue(normalizedSourcePath, out targetDoc))
                                {
                                    matchedTargetPath = targetDoc.RelativePath;
                                }
                            }

                            if (targetDoc != null && matchedTargetPath != null)
                            {
                                comparison.TargetItemId = targetDoc.Id;
                                comparison.TargetSizeBytes = targetDoc.SizeBytes;
                                comparison.TargetVersionCount = targetDoc.VersionCount;
                                comparison.TargetAbsolutePath = BuildAbsoluteUrl(pair.TargetUrl, targetDoc.ServerRelativeUrl);
                                comparison.TargetCreated = targetDoc.Created;
                                comparison.TargetModified = targetDoc.Modified;

                                // Check for concerning size issues (only for files, not folders):
                                // 1. Target is 0 bytes when source > 0
                                // 2. Source > 50KB and target < 30% of source size
                                bool hasSizeIssue = false;
                                if (sourceDoc.ItemType == DocumentCompareItemType.File)
                                {
                                    if (targetDoc.SizeBytes == 0 && sourceDoc.SizeBytes > 0)
                                    {
                                        hasSizeIssue = true;
                                    }
                                    else if (sourceDoc.SizeBytes > 50 * 1024 && targetDoc.SizeBytes < sourceDoc.SizeBytes * 0.3)
                                    {
                                        hasSizeIssue = true;
                                    }
                                }

                                if (hasSizeIssue)
                                {
                                    comparison.Status = DocumentCompareStatus.SizeIssue;
                                    sizeIssueCount++;
                                }
                                else
                                {
                                    comparison.Status = DocumentCompareStatus.Found;
                                    foundCount++;
                                }

                                // Remove from target dicts to track target-only later
                                targetByPath.Remove(matchedTargetPath);
                                if (config.UseShareGateNormalization && targetByNormalizedPath != null)
                                {
                                    targetByNormalizedPath.Remove(NormalizePathForShareGate(matchedTargetPath));
                                }
                            }
                            else
                            {
                                comparison.Status = DocumentCompareStatus.SourceOnly;
                                sourceOnlyCount++;
                            }

                            siteResult.DocumentComparisons.Add(comparison);
                        }

                        // Add target-only documents and folders
                        int targetOnlyCount = 0;
                        foreach (var targetDoc in targetByPath.Values)
                        {
                            siteResult.DocumentComparisons.Add(new DocumentCompareItem
                            {
                                SourceSiteUrl = pair.SourceUrl,
                                TargetSiteUrl = pair.TargetUrl,
                                LibraryName = sourceLibrary.Title,
                                ItemType = targetDoc.ItemType,
                                FileName = targetDoc.FileName,
                                FileExtension = targetDoc.ItemType == DocumentCompareItemType.Folder ? "" : Path.GetExtension(targetDoc.FileName).TrimStart('.'),
                                RelativePath = targetDoc.RelativePath,
                                TargetItemId = targetDoc.Id,
                                TargetSizeBytes = targetDoc.SizeBytes,
                                TargetVersionCount = targetDoc.VersionCount,
                                TargetAbsolutePath = BuildAbsoluteUrl(pair.TargetUrl, targetDoc.ServerRelativeUrl),
                                TargetCreated = targetDoc.Created,
                                TargetModified = targetDoc.Modified,
                                Status = DocumentCompareStatus.TargetOnly
                            });
                            targetOnlyCount++;
                        }

                        var newerAtSourceCount = siteResult.DocumentComparisons.Count(dc => dc.IsNewerAtSource);
                        result.Log($"    Results: {foundCount} found, {sizeIssueCount} size issues, {sourceOnlyCount} source-only, {targetOnlyCount} target-only, {newerAtSourceCount} newer-at-source");
                    }

                    siteResult.Success = true;
                    result.SuccessfulPairs++;

                    result.Log($"  Site completed: {siteResult.TotalDocuments} documents compared");
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

                // Report completed site for real-time UI update
                progress?.Report(new TaskProgress
                {
                    CurrentSite = processedCount,
                    TotalSites = config.SitePairs.Count,
                    CurrentSiteUrl = pair.SourceUrl,
                    Message = $"Completed {processedCount}/{config.SitePairs.Count}: {pair.SourceUrl}",
                    CompletedSiteResult = siteResult
                });
            }

            result.CompletedAt = DateTime.UtcNow;
            result.Success = result.FailedPairs == 0;

            var (found, sizeIssues, sourceOnly, targetOnly, newerAtSource) = result.GetSummary();
            result.Log($"Task completed. Successful: {result.SuccessfulPairs}, Failed: {result.FailedPairs}");
            result.Log($"Total: {found} found, {sizeIssues} size issues, {sourceOnly} source-only, {targetOnly} target-only, {newerAtSource} newer-at-source");

            task.Status = result.FailedPairs == 0 ? Models.TaskStatus.Completed : Models.TaskStatus.Failed;
            task.CompletedAt = DateTime.UtcNow;

            if (result.FailedPairs > 0)
            {
                task.LastError = $"{result.FailedPairs} site pair(s) failed";
            }

            } // end inner try
            finally
            {
                sourceService.Dispose();
                targetService.Dispose();
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
        await SaveDocumentCompareResultAsync(result);

        return result;
    }

    private static string BuildAbsoluteUrl(string siteUrl, string serverRelativeUrl)
    {
        try
        {
            var siteUri = new Uri(siteUrl);
            return $"{siteUri.Scheme}://{siteUri.Host}{serverRelativeUrl}";
        }
        catch
        {
            return serverRelativeUrl;
        }
    }

    private static bool IsAuthError(string? errorMessage)
    {
        if (string.IsNullOrEmpty(errorMessage)) return false;
        return errorMessage.Contains("Authentication failed", StringComparison.OrdinalIgnoreCase)
            || errorMessage.Contains("Access denied", StringComparison.OrdinalIgnoreCase)
            || errorMessage.Contains("cookies may have expired", StringComparison.OrdinalIgnoreCase);
    }

    #region Document Compare Cache

    /// <summary>
    /// Gets the cache file path for a specific site and library.
    /// </summary>
    private string GetCacheFilePath(string siteUrl, string libraryTitle)
    {
        // Create a safe filename from site URL and library
        var key = $"{siteUrl}|{libraryTitle}".ToLowerInvariant();
        var hash = Convert.ToBase64String(
            System.Security.Cryptography.SHA256.HashData(
                System.Text.Encoding.UTF8.GetBytes(key)))
            .Replace("/", "_")
            .Replace("+", "-")
            .Replace("=", "")[..16];
        return Path.Combine(_cacheFolder, $"{hash}.json");
    }

    /// <summary>
    /// Tries to get cached documents for a site/library combination.
    /// </summary>
    private DocumentCompareCacheEntry? TryGetCachedDocuments(string siteUrl, string libraryTitle, int expirationHours)
    {
        var filePath = GetCacheFilePath(siteUrl, libraryTitle);
        if (!File.Exists(filePath))
            return null;

        try
        {
            var json = File.ReadAllText(filePath);
            var entry = JsonSerializer.Deserialize<DocumentCompareCacheEntry>(json, _jsonOptions);
            if (entry != null && entry.IsValid(expirationHours))
            {
                return entry;
            }
            // Cache expired, delete it
            File.Delete(filePath);
        }
        catch
        {
            // Ignore cache read errors
        }
        return null;
    }

    /// <summary>
    /// Saves documents to cache.
    /// </summary>
    private void SaveToCache(string siteUrl, string libraryTitle, List<DocumentCompareSourceItem> documents)
    {
        var entry = new DocumentCompareCacheEntry
        {
            CachedAt = DateTime.UtcNow,
            SiteUrl = siteUrl,
            LibraryTitle = libraryTitle,
            Documents = documents
        };

        var filePath = GetCacheFilePath(siteUrl, libraryTitle);
        try
        {
            var json = JsonSerializer.Serialize(entry, _jsonOptions);
            File.WriteAllText(filePath, json);
        }
        catch
        {
            // Ignore cache write errors
        }
    }

    #endregion

    /// <summary>
    /// Normalizes a path by simulating ShareGate's character replacement during migration.
    /// ShareGate URL-encodes filenames first (space → %20, etc.), then replaces % with _.
    /// It also replaces special characters: " * : &lt; &gt; ? \ &amp; # % { } ~ with underscore.
    /// And handles consecutive dots: .. → _., ... → __., etc.
    /// Note: Forward slash is preserved as path separator.
    /// </summary>
    private static string NormalizePathForShareGate(string path)
    {
        if (string.IsNullOrEmpty(path))
            return path;

        // Step 1: Replace special characters with underscores
        // " * : < > ? \ & # % { } ~
        var result = new System.Text.StringBuilder(path.Length);

        foreach (var c in path)
        {
            if (c == '/' || c == '.')
            {
                // Preserve path separators and dots (for extensions)
                result.Append(c);
            }
            else if (c == ' ')
            {
                // ShareGate URL-encodes spaces to %20, then replaces % with _
                // Net effect: space becomes _20
                result.Append("_20");
            }
            else if (c == '"' || c == '*' || c == ':' || c == '<' || c == '>' ||
                     c == '?' || c == '\\' || c == '&' || c == '#' || c == '%' ||
                     c == '{' || c == '}' || c == '~')
            {
                result.Append('_');
            }
            else
            {
                result.Append(c);
            }
        }

        // Step 2: Handle consecutive dots: ShareGate replaces dot-pairs with _.
        // .. → _.   (1 pair)
        // ... → __.  (pair + remaining dot, needs 2 passes)
        // .... → _._. (2 pairs)
        // ...... → _._._. (3 pairs)
        var normalized = result.ToString();
        while (normalized.Contains(".."))
        {
            normalized = normalized.Replace("..", "_.");
        }

        return normalized;
    }

    public async Task<List<DocumentCompareResult>> GetDocumentCompareResultsAsync(Guid taskId)
    {
        var results = new List<DocumentCompareResult>();
        var pattern = $"doccompare_{taskId}_*.json";
        var files = Directory.GetFiles(_resultsFolder, pattern)
            .OrderByDescending(f => f);

        foreach (var file in files)
        {
            try
            {
                // Use streaming deserialization to avoid huge string allocation for large files
                await using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read, 65536, useAsync: true);
                var result = await JsonSerializer.DeserializeAsync<DocumentCompareResult>(stream, _jsonOptions);
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

    public async Task<DocumentCompareResult?> GetLatestDocumentCompareResultAsync(Guid taskId)
    {
        var pattern = $"doccompare_{taskId}_*.json";
        var latestFile = Directory.GetFiles(_resultsFolder, pattern)
            .OrderByDescending(f => f)
            .FirstOrDefault();

        if (latestFile == null)
            return null;

        try
        {
            // Use streaming deserialization to avoid huge string allocation for large files
            await using var stream = new FileStream(latestFile, FileMode.Open, FileAccess.Read, FileShare.Read, 65536, useAsync: true);
            return await JsonSerializer.DeserializeAsync<DocumentCompareResult>(stream, _jsonOptions);
        }
        catch
        {
            return null;
        }
    }

    public async Task SaveDocumentCompareResultAsync(DocumentCompareResult result)
    {
        var timestamp = result.ExecutedAt.ToString("yyyyMMdd_HHmmss");
        var fileName = $"doccompare_{result.TaskId}_{timestamp}.json";
        var filePath = Path.Combine(_resultsFolder, fileName);

        // Use streaming serialization to avoid huge string allocation
        await using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 65536, useAsync: true);
        await JsonSerializer.SerializeAsync(stream, result, _jsonOptions);
    }

    public async Task<SiteAccessResult> ExecuteSiteAccessCheckAsync(
        TaskDefinition task,
        IAuthenticationService authService,
        IConnectionManager connectionManager,
        IProgress<TaskProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var result = new SiteAccessResult
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
            await SaveSiteAccessResultAsync(result);
            return result;
        }

        SiteAccessConfiguration config;
        try
        {
            config = JsonSerializer.Deserialize<SiteAccessConfiguration>(task.ConfigurationJson, _jsonOptions)
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
            await SaveSiteAccessResultAsync(result);
            return result;
        }

        result.Log($"Starting site access check for {config.SitePairs.Count} site pairs");

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

            // Set account names from cookies (actual user emails)
            result.SourceAccount = sourceCookies.UserEmail ?? sourceConnection.Name;
            result.TargetAccount = targetCookies.UserEmail ?? targetConnection.Name;
            result.Log($"Source account: {result.SourceAccount}");
            result.Log($"Target account: {result.TargetAccount}");

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
                    Message = $"Checking {processedCount}/{config.SitePairs.Count}: {pair.SourceUrl}"
                });

                result.Log($"Checking pair {processedCount}: {pair.SourceUrl} <-> {pair.TargetUrl}");

                var pairResult = new SitePairAccessResult
                {
                    SourceSiteUrl = pair.SourceUrl,
                    TargetSiteUrl = pair.TargetUrl
                };

                // Check source site access
                pairResult.SourceResult = await CheckSiteAccessAsync(
                    sourceService, pair.SourceUrl, result.SourceAccount ?? sourceConnection.Name, true);
                result.Log($"  Source: {pairResult.SourceResult.StatusDescription}");

                // Check target site access
                pairResult.TargetResult = await CheckSiteAccessAsync(
                    targetService, pair.TargetUrl, result.TargetAccount ?? targetConnection.Name, false);
                result.Log($"  Target: {pairResult.TargetResult.StatusDescription}");

                result.PairResults.Add(pairResult);
                result.TotalPairsProcessed++;

                // Report completed pair for real-time UI update
                progress?.Report(new TaskProgress
                {
                    CurrentSite = processedCount,
                    TotalSites = config.SitePairs.Count,
                    CurrentSiteUrl = pair.SourceUrl,
                    Message = $"Checked {processedCount}/{config.SitePairs.Count}: {pair.SourceUrl}",
                    CompletedAccessPairResult = pairResult
                });
            }

            result.CompletedAt = DateTime.UtcNow;
            result.Success = true;

            result.Log($"Task completed. Processed: {result.TotalPairsProcessed} pairs");
            result.Log($"Source: {result.SourceAccessibleCount} accessible, {result.SourceAccessDeniedCount} access denied, {result.SourceOtherIssuesCount} other issues");
            result.Log($"Target: {result.TargetAccessibleCount} accessible, {result.TargetAccessDeniedCount} access denied, {result.TargetOtherIssuesCount} other issues");

            task.Status = Models.TaskStatus.Completed;
            task.CompletedAt = DateTime.UtcNow;
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
        await SaveSiteAccessResultAsync(result);

        return result;
    }

    private static async Task<SiteAccessCheckItem> CheckSiteAccessAsync(
        SharePointService service, string siteUrl, string accountName, bool isSource)
    {
        var result = new SiteAccessCheckItem
        {
            SiteUrl = siteUrl,
            AccountUsed = accountName,
            IsSource = isSource
        };

        try
        {
            var siteInfo = await service.GetSiteInfoAsync(siteUrl);

            if (siteInfo.IsConnected)
            {
                result.SiteTitle = siteInfo.Title ?? "";
                result.Status = SiteAccessStatus.Accessible;
            }
            else
            {
                result.Status = ParseAccessStatus(siteInfo.ErrorMessage);
                result.ErrorMessage = siteInfo.ErrorMessage;
            }
        }
        catch (HttpRequestException ex)
        {
            result.Status = ex.StatusCode switch
            {
                System.Net.HttpStatusCode.Forbidden => SiteAccessStatus.AccessDenied,
                System.Net.HttpStatusCode.NotFound => SiteAccessStatus.NotFound,
                System.Net.HttpStatusCode.Unauthorized => SiteAccessStatus.AuthenticationRequired,
                _ => SiteAccessStatus.Error
            };
            result.ErrorMessage = ex.Message;
        }
        catch (Exception ex)
        {
            result.Status = SiteAccessStatus.Error;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    private static SiteAccessStatus ParseAccessStatus(string? errorMessage)
    {
        if (string.IsNullOrEmpty(errorMessage))
            return SiteAccessStatus.Error;

        if (errorMessage.Contains("403") || errorMessage.Contains("Forbidden", StringComparison.OrdinalIgnoreCase))
            return SiteAccessStatus.AccessDenied;

        if (errorMessage.Contains("404") || errorMessage.Contains("Not Found", StringComparison.OrdinalIgnoreCase))
            return SiteAccessStatus.NotFound;

        if (errorMessage.Contains("401") || errorMessage.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase))
            return SiteAccessStatus.AuthenticationRequired;

        return SiteAccessStatus.Error;
    }

    public async Task<List<SiteAccessResult>> GetSiteAccessResultsAsync(Guid taskId)
    {
        var results = new List<SiteAccessResult>();
        var pattern = $"siteaccess_{taskId}_*.json";
        var files = Directory.GetFiles(_resultsFolder, pattern)
            .OrderByDescending(f => f);

        foreach (var file in files)
        {
            try
            {
                var json = await File.ReadAllTextAsync(file);
                var result = JsonSerializer.Deserialize<SiteAccessResult>(json, _jsonOptions);
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

    public async Task<SiteAccessResult?> GetLatestSiteAccessResultAsync(Guid taskId)
    {
        var results = await GetSiteAccessResultsAsync(taskId);
        return results.FirstOrDefault();
    }

    public async Task SaveSiteAccessResultAsync(SiteAccessResult result)
    {
        var timestamp = result.ExecutedAt.ToString("yyyyMMdd_HHmmss");
        var fileName = $"siteaccess_{result.TaskId}_{timestamp}.json";
        var filePath = Path.Combine(_resultsFolder, fileName);

        var json = JsonSerializer.Serialize(result, _jsonOptions);
        await File.WriteAllTextAsync(filePath, json);
    }
}

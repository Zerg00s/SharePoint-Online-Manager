namespace SharePointOnlineManager.Models;

/// <summary>
/// Context data passed when navigating to the task type selection screen.
/// </summary>
public class TaskCreationContext
{
    public required Connection Connection { get; init; }
    public required List<SiteCollection> SelectedSites { get; init; }
}

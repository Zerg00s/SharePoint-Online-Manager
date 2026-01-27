namespace SharePointOnlineManager.Navigation;

/// <summary>
/// Interface for screen navigation within the application.
/// </summary>
public interface INavigationService
{
    /// <summary>
    /// Gets the current screen.
    /// </summary>
    BaseScreen? CurrentScreen { get; }

    /// <summary>
    /// Gets whether there are screens in the navigation stack to go back to.
    /// </summary>
    bool CanGoBack { get; }

    /// <summary>
    /// Navigates to a screen of the specified type.
    /// </summary>
    Task NavigateToAsync<TScreen>(object? parameter = null) where TScreen : BaseScreen;

    /// <summary>
    /// Navigates to the specified screen instance.
    /// </summary>
    Task NavigateToAsync(BaseScreen screen, object? parameter = null);

    /// <summary>
    /// Navigates back to the previous screen.
    /// </summary>
    Task<bool> GoBackAsync();

    /// <summary>
    /// Navigates to the home screen, clearing the navigation stack.
    /// </summary>
    Task NavigateToHomeAsync();

    /// <summary>
    /// Sets the status bar message.
    /// </summary>
    void SetStatus(string message);

    /// <summary>
    /// Updates the title bar with the current screen's title.
    /// </summary>
    void UpdateTitle();

    /// <summary>
    /// Shows a loading indicator with a message.
    /// </summary>
    void ShowLoading(string message = "Loading...");

    /// <summary>
    /// Hides the loading indicator.
    /// </summary>
    void HideLoading();

    /// <summary>
    /// Event raised when navigation occurs.
    /// </summary>
    event EventHandler<NavigationEventArgs>? Navigated;
}

/// <summary>
/// Event arguments for navigation events.
/// </summary>
public class NavigationEventArgs : EventArgs
{
    public BaseScreen? PreviousScreen { get; init; }
    public BaseScreen CurrentScreen { get; init; } = null!;
    public object? Parameter { get; init; }
    public bool IsBackNavigation { get; init; }
}

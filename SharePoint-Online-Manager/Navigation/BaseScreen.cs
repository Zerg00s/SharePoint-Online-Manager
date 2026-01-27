namespace SharePointOnlineManager.Navigation;

/// <summary>
/// Abstract base class for all screens in the application.
/// Screens are UserControls that can be navigated to and from.
/// </summary>
public abstract class BaseScreen : UserControl
{
    /// <summary>
    /// Gets or sets the navigation service for this screen.
    /// </summary>
    protected INavigationService? NavigationService { get; private set; }

    /// <summary>
    /// Gets or sets the service provider for dependency resolution.
    /// </summary>
    protected IServiceProvider? ServiceProvider { get; private set; }

    /// <summary>
    /// Gets the display title for this screen (shown in breadcrumb/title bar).
    /// </summary>
    public abstract string ScreenTitle { get; }

    /// <summary>
    /// Gets whether the back button should be visible for this screen.
    /// </summary>
    public virtual bool ShowBackButton => true;

    /// <summary>
    /// Initializes the screen with navigation service and service provider.
    /// </summary>
    public void Initialize(INavigationService navigationService, IServiceProvider serviceProvider)
    {
        NavigationService = navigationService;
        ServiceProvider = serviceProvider;
        OnInitialize();
    }

    /// <summary>
    /// Called when the screen is initialized. Override to perform setup.
    /// </summary>
    protected virtual void OnInitialize()
    {
    }

    /// <summary>
    /// Called when the screen is navigated to (becomes visible).
    /// </summary>
    public virtual Task OnNavigatedToAsync(object? parameter = null)
    {
        return Task.CompletedTask;
    }

    /// <summary>
    /// Called when navigating away from this screen.
    /// Return false to cancel navigation.
    /// </summary>
    public virtual Task<bool> OnNavigatingFromAsync()
    {
        return Task.FromResult(true);
    }

    /// <summary>
    /// Called when the screen is navigated from (becomes hidden).
    /// </summary>
    public virtual void OnNavigatedFrom()
    {
    }

    /// <summary>
    /// Updates the status bar message.
    /// </summary>
    protected void SetStatus(string message)
    {
        NavigationService?.SetStatus(message);
    }

    /// <summary>
    /// Shows a loading indicator.
    /// </summary>
    protected void ShowLoading(string message = "Loading...")
    {
        NavigationService?.ShowLoading(message);
    }

    /// <summary>
    /// Hides the loading indicator.
    /// </summary>
    protected void HideLoading()
    {
        NavigationService?.HideLoading();
    }

    /// <summary>
    /// Updates the title bar with this screen's title.
    /// </summary>
    protected void UpdateTitle()
    {
        NavigationService?.UpdateTitle();
    }

    /// <summary>
    /// Helper method to resolve a service from the service provider.
    /// </summary>
    protected T? GetService<T>() where T : class
    {
        return ServiceProvider?.GetService(typeof(T)) as T;
    }

    /// <summary>
    /// Helper method to resolve a required service from the service provider.
    /// </summary>
    protected T GetRequiredService<T>() where T : class
    {
        return (T)(ServiceProvider?.GetService(typeof(T))
            ?? throw new InvalidOperationException($"Service {typeof(T).Name} not registered."));
    }
}

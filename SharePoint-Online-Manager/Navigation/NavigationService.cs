namespace SharePointOnlineManager.Navigation;

/// <summary>
/// Panel-swapping implementation of navigation service.
/// </summary>
public class NavigationService : INavigationService
{
    private readonly Panel _contentPanel;
    private readonly IServiceProvider _serviceProvider;
    private readonly Stack<BaseScreen> _navigationStack = new();
    private readonly Action<string> _statusUpdater;
    private readonly Action<string> _loadingShower;
    private readonly Action _loadingHider;
    private readonly Action<string> _titleUpdater;
    private readonly Action<bool> _backButtonUpdater;

    public BaseScreen? CurrentScreen => _navigationStack.Count > 0 ? _navigationStack.Peek() : null;
    public bool CanGoBack => _navigationStack.Count > 1;

    public event EventHandler<NavigationEventArgs>? Navigated;

    public NavigationService(
        Panel contentPanel,
        IServiceProvider serviceProvider,
        Action<string> statusUpdater,
        Action<string> loadingShower,
        Action loadingHider,
        Action<string> titleUpdater,
        Action<bool> backButtonUpdater)
    {
        _contentPanel = contentPanel;
        _serviceProvider = serviceProvider;
        _statusUpdater = statusUpdater;
        _loadingShower = loadingShower;
        _loadingHider = loadingHider;
        _titleUpdater = titleUpdater;
        _backButtonUpdater = backButtonUpdater;
    }

    public async Task NavigateToAsync<TScreen>(object? parameter = null) where TScreen : BaseScreen
    {
        var screen = (TScreen)(Activator.CreateInstance(typeof(TScreen))
            ?? throw new InvalidOperationException($"Could not create instance of {typeof(TScreen).Name}"));

        await NavigateToAsync(screen, parameter);
    }

    public async Task NavigateToAsync(BaseScreen screen, object? parameter = null)
    {
        var previousScreen = CurrentScreen;

        // Check if we can navigate away from current screen
        if (previousScreen != null)
        {
            var canNavigate = await previousScreen.OnNavigatingFromAsync();
            if (!canNavigate)
            {
                return;
            }
        }

        // Initialize the new screen
        screen.Initialize(this, _serviceProvider);
        screen.Dock = DockStyle.Fill;

        // Update the content panel
        _contentPanel.SuspendLayout();
        try
        {
            // Hide and notify previous screen
            if (previousScreen != null)
            {
                previousScreen.Visible = false;
                previousScreen.OnNavigatedFrom();
            }

            // Add new screen if not already in panel
            if (!_contentPanel.Controls.Contains(screen))
            {
                _contentPanel.Controls.Add(screen);
            }

            // Push to navigation stack
            _navigationStack.Push(screen);

            // Show new screen
            screen.Visible = true;
            screen.BringToFront();
        }
        finally
        {
            _contentPanel.ResumeLayout(true);
        }

        // Update UI elements
        _titleUpdater(screen.ScreenTitle);
        _backButtonUpdater(CanGoBack && screen.ShowBackButton);
        SetStatus("Ready");

        // Notify the screen
        await screen.OnNavigatedToAsync(parameter);

        // Raise event
        Navigated?.Invoke(this, new NavigationEventArgs
        {
            PreviousScreen = previousScreen,
            CurrentScreen = screen,
            Parameter = parameter,
            IsBackNavigation = false
        });
    }

    public async Task<bool> GoBackAsync()
    {
        if (!CanGoBack)
        {
            return false;
        }

        var currentScreen = _navigationStack.Pop();

        // Check if we can navigate away
        var canNavigate = await currentScreen.OnNavigatingFromAsync();
        if (!canNavigate)
        {
            // Push it back if navigation was cancelled
            _navigationStack.Push(currentScreen);
            return false;
        }

        var previousScreen = CurrentScreen!;

        // Update the content panel
        _contentPanel.SuspendLayout();
        try
        {
            // Hide current screen
            currentScreen.Visible = false;
            currentScreen.OnNavigatedFrom();

            // Remove current screen from panel
            _contentPanel.Controls.Remove(currentScreen);
            currentScreen.Dispose();

            // Show previous screen
            previousScreen.Visible = true;
            previousScreen.BringToFront();
        }
        finally
        {
            _contentPanel.ResumeLayout(true);
        }

        // Update UI elements
        _titleUpdater(previousScreen.ScreenTitle);
        _backButtonUpdater(CanGoBack && previousScreen.ShowBackButton);
        SetStatus("Ready");

        // Notify the screen
        await previousScreen.OnNavigatedToAsync();

        // Raise event
        Navigated?.Invoke(this, new NavigationEventArgs
        {
            PreviousScreen = currentScreen,
            CurrentScreen = previousScreen,
            IsBackNavigation = true
        });

        return true;
    }

    public async Task NavigateToHomeAsync()
    {
        // Pop all screens except the home screen
        while (_navigationStack.Count > 1)
        {
            var screen = _navigationStack.Pop();
            screen.OnNavigatedFrom();
            _contentPanel.Controls.Remove(screen);
            screen.Dispose();
        }

        if (CurrentScreen != null)
        {
            _contentPanel.SuspendLayout();
            try
            {
                CurrentScreen.Visible = true;
                CurrentScreen.BringToFront();
            }
            finally
            {
                _contentPanel.ResumeLayout(true);
            }

            _titleUpdater(CurrentScreen.ScreenTitle);
            _backButtonUpdater(false);
            SetStatus("Ready");

            await CurrentScreen.OnNavigatedToAsync();

            Navigated?.Invoke(this, new NavigationEventArgs
            {
                CurrentScreen = CurrentScreen,
                IsBackNavigation = true
            });
        }
    }

    public void SetStatus(string message)
    {
        _statusUpdater(message);
    }

    public void ShowLoading(string message = "Loading...")
    {
        _loadingShower(message);
    }

    public void HideLoading()
    {
        _loadingHider();
    }

    public void UpdateTitle()
    {
        if (CurrentScreen != null)
        {
            _titleUpdater(CurrentScreen.ScreenTitle);
        }
    }
}

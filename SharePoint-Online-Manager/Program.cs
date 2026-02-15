using SharePointOnlineManager.Forms;
using SharePointOnlineManager.Services;

namespace SharePointOnlineManager;

static class Program
{
    /// <summary>
    /// The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        // Create service provider
        var serviceProvider = new ServiceProvider();

        Application.Run(new MainForm(serviceProvider));
    }
}

/// <summary>
/// Simple service provider for dependency injection.
/// </summary>
public class ServiceProvider : IServiceProvider
{
    private readonly Dictionary<Type, object> _services = new();
    private readonly Dictionary<Type, Func<object>> _factories = new();

    public ServiceProvider()
    {
        // Register singleton services
        var authService = new AuthenticationService();
        _services[typeof(IAuthenticationService)] = authService;
        _services[typeof(AuthenticationService)] = authService;

        var connectionManager = new ConnectionManager(authService);
        _services[typeof(IConnectionManager)] = connectionManager;
        _services[typeof(ConnectionManager)] = connectionManager;

        var taskService = new TaskService();
        _services[typeof(ITaskService)] = taskService;
        _services[typeof(TaskService)] = taskService;

        var csvExporter = new CsvExporter();
        _services[typeof(CsvExporter)] = csvExporter;

        var excelExporter = new ExcelExporter();
        _services[typeof(ExcelExporter)] = excelExporter;

        var tenantPairService = new TenantPairService();
        _services[typeof(ITenantPairService)] = tenantPairService;
        _services[typeof(TenantPairService)] = tenantPairService;
    }

    public object? GetService(Type serviceType)
    {
        if (_services.TryGetValue(serviceType, out var service))
        {
            return service;
        }

        if (_factories.TryGetValue(serviceType, out var factory))
        {
            return factory();
        }

        return null;
    }
}

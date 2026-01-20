using SharePointOnlineManager.Data;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Manages SharePoint connection configurations with JSON persistence.
/// </summary>
public class ConnectionManager : IConnectionManager
{
    private readonly IDataStore<Connection> _dataStore;
    private readonly IAuthenticationService _authService;

    public ConnectionManager(IAuthenticationService authService)
    {
        _authService = authService;
        _dataStore = new JsonDataStore<Connection>(
            "connections.json",
            c => c.Id,
            (c, id) => c.Id = id);
    }

    public async Task<List<Connection>> GetAllConnectionsAsync()
    {
        var connections = await _dataStore.GetAllAsync();
        return connections.OrderByDescending(c => c.LastConnectedAt ?? c.CreatedAt).ToList();
    }

    public async Task<Connection?> GetConnectionAsync(Guid id)
    {
        return await _dataStore.GetByIdAsync(id);
    }

    public async Task SaveConnectionAsync(Connection connection)
    {
        if (connection.Id == Guid.Empty)
        {
            connection.Id = Guid.NewGuid();
            connection.CreatedAt = DateTime.UtcNow;
        }

        await _dataStore.SaveAsync(connection);
    }

    public async Task DeleteConnectionAsync(Guid id)
    {
        var connection = await _dataStore.GetByIdAsync(id);
        if (connection != null)
        {
            // Clear any stored credentials for this connection
            ClearCredentials(connection);
        }

        await _dataStore.DeleteAsync(id);
    }

    public async Task UpdateLastConnectedAsync(Guid connectionId)
    {
        var connection = await _dataStore.GetByIdAsync(connectionId);
        if (connection != null)
        {
            connection.LastConnectedAt = DateTime.UtcNow;
            await _dataStore.SaveAsync(connection);
        }
    }

    public bool HasStoredCredentials(Connection connection)
    {
        return _authService.HasStoredCredentials(connection.CookieDomain);
    }

    public void ClearCredentials(Connection connection)
    {
        _authService.ClearCredentials(connection.CookieDomain);
    }
}

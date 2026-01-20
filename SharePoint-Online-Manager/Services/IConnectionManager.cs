using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Interface for managing SharePoint connection configurations.
/// </summary>
public interface IConnectionManager
{
    /// <summary>
    /// Gets all saved connections.
    /// </summary>
    Task<List<Connection>> GetAllConnectionsAsync();

    /// <summary>
    /// Gets a connection by its ID.
    /// </summary>
    Task<Connection?> GetConnectionAsync(Guid id);

    /// <summary>
    /// Saves a connection (insert or update).
    /// </summary>
    Task SaveConnectionAsync(Connection connection);

    /// <summary>
    /// Deletes a connection by its ID.
    /// </summary>
    Task DeleteConnectionAsync(Guid id);

    /// <summary>
    /// Updates the last connected timestamp for a connection.
    /// </summary>
    Task UpdateLastConnectedAsync(Guid connectionId);

    /// <summary>
    /// Gets whether a connection has stored credentials.
    /// </summary>
    bool HasStoredCredentials(Connection connection);

    /// <summary>
    /// Clears stored credentials for a connection.
    /// </summary>
    void ClearCredentials(Connection connection);
}

namespace SharePointOnlineManager.Data;

/// <summary>
/// Generic interface for data persistence operations.
/// </summary>
/// <typeparam name="T">The type of entity to store.</typeparam>
public interface IDataStore<T> where T : class
{
    /// <summary>
    /// Gets all stored items.
    /// </summary>
    Task<List<T>> GetAllAsync();

    /// <summary>
    /// Gets an item by its identifier.
    /// </summary>
    Task<T?> GetByIdAsync(Guid id);

    /// <summary>
    /// Saves an item (insert or update).
    /// </summary>
    Task SaveAsync(T item);

    /// <summary>
    /// Deletes an item by its identifier.
    /// </summary>
    Task DeleteAsync(Guid id);

    /// <summary>
    /// Saves all items, replacing the existing collection.
    /// </summary>
    Task SaveAllAsync(List<T> items);
}

using System.Text.Json;
using System.Text.Json.Serialization;

namespace SharePointOnlineManager.Data;

/// <summary>
/// JSON file-based data store implementation.
/// </summary>
/// <typeparam name="T">The type of entity to store. Must have a Guid Id property.</typeparam>
public class JsonDataStore<T> : IDataStore<T> where T : class
{
    private readonly string _filePath;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private readonly JsonSerializerOptions _jsonOptions;
    private readonly Func<T, Guid> _idSelector;
    private readonly Action<T, Guid> _idSetter;

    /// <summary>
    /// Creates a new JsonDataStore instance.
    /// </summary>
    /// <param name="fileName">The JSON file name (e.g., "connections.json").</param>
    /// <param name="idSelector">Function to get the Id from an entity.</param>
    /// <param name="idSetter">Action to set the Id on an entity.</param>
    public JsonDataStore(string fileName, Func<T, Guid> idSelector, Action<T, Guid> idSetter)
    {
        _idSelector = idSelector;
        _idSetter = idSetter;

        var appDataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SharePointOnlineManager",
            "data");

        Directory.CreateDirectory(appDataPath);
        _filePath = Path.Combine(appDataPath, fileName);

        _jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Converters = { new JsonStringEnumConverter() }
        };
    }

    public async Task<List<T>> GetAllAsync()
    {
        await _lock.WaitAsync();
        try
        {
            if (!File.Exists(_filePath))
            {
                return [];
            }

            var json = await File.ReadAllTextAsync(_filePath);
            if (string.IsNullOrWhiteSpace(json))
            {
                return [];
            }

            return JsonSerializer.Deserialize<List<T>>(json, _jsonOptions) ?? [];
        }
        catch (JsonException)
        {
            return [];
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task<T?> GetByIdAsync(Guid id)
    {
        var items = await GetAllAsync();
        return items.FirstOrDefault(item => _idSelector(item) == id);
    }

    public async Task SaveAsync(T item)
    {
        await _lock.WaitAsync();
        try
        {
            var items = await LoadItemsUnsafeAsync();
            var id = _idSelector(item);

            if (id == Guid.Empty)
            {
                _idSetter(item, Guid.NewGuid());
                id = _idSelector(item);
            }

            var existingIndex = items.FindIndex(i => _idSelector(i) == id);
            if (existingIndex >= 0)
            {
                items[existingIndex] = item;
            }
            else
            {
                items.Add(item);
            }

            await SaveItemsUnsafeAsync(items);
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task DeleteAsync(Guid id)
    {
        await _lock.WaitAsync();
        try
        {
            var items = await LoadItemsUnsafeAsync();
            var removed = items.RemoveAll(item => _idSelector(item) == id);

            if (removed > 0)
            {
                await SaveItemsUnsafeAsync(items);
            }
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task SaveAllAsync(List<T> items)
    {
        await _lock.WaitAsync();
        try
        {
            await SaveItemsUnsafeAsync(items);
        }
        finally
        {
            _lock.Release();
        }
    }

    private async Task<List<T>> LoadItemsUnsafeAsync()
    {
        if (!File.Exists(_filePath))
        {
            return [];
        }

        var json = await File.ReadAllTextAsync(_filePath);
        if (string.IsNullOrWhiteSpace(json))
        {
            return [];
        }

        try
        {
            return JsonSerializer.Deserialize<List<T>>(json, _jsonOptions) ?? [];
        }
        catch (JsonException)
        {
            return [];
        }
    }

    private async Task SaveItemsUnsafeAsync(List<T> items)
    {
        var directory = Path.GetDirectoryName(_filePath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(items, _jsonOptions);
        await File.WriteAllTextAsync(_filePath, json);
    }
}

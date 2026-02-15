using System.Text.Json;
using SharePointOnlineManager.Models;

namespace SharePointOnlineManager.Services;

/// <summary>
/// Interface for managing tenant pairs.
/// </summary>
public interface ITenantPairService
{
    Task<List<TenantPair>> GetAllPairsAsync();
    Task<TenantPair?> GetPairAsync(Guid id);
    Task<TenantPair?> GetPairByConnectionsAsync(Guid sourceConnectionId, Guid targetConnectionId);
    Task SavePairAsync(TenantPair pair);
    Task DeletePairAsync(Guid id);
}

/// <summary>
/// Service for managing tenant pair configurations.
/// </summary>
public class TenantPairService : ITenantPairService
{
    private static readonly string ConfigFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SharePointOnlineManager");

    private static readonly string ConfigFile = Path.Combine(ConfigFolder, "tenant_pairs.json");

    private TenantPairConfiguration? _config;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public async Task<List<TenantPair>> GetAllPairsAsync()
    {
        await EnsureLoadedAsync();
        return _config!.Pairs.ToList();
    }

    public async Task<TenantPair?> GetPairAsync(Guid id)
    {
        await EnsureLoadedAsync();
        return _config!.Pairs.FirstOrDefault(p => p.Id == id);
    }

    public async Task<TenantPair?> GetPairByConnectionsAsync(Guid sourceConnectionId, Guid targetConnectionId)
    {
        await EnsureLoadedAsync();
        return _config!.Pairs.FirstOrDefault(p =>
            p.SourceConnectionId == sourceConnectionId &&
            p.TargetConnectionId == targetConnectionId);
    }

    public async Task SavePairAsync(TenantPair pair)
    {
        await _lock.WaitAsync();
        try
        {
            await EnsureLoadedAsync();

            var existing = _config!.Pairs.FirstOrDefault(p => p.Id == pair.Id);
            if (existing != null)
            {
                _config.Pairs.Remove(existing);
            }

            _config.Pairs.Add(pair);
            await SaveConfigAsync();
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task DeletePairAsync(Guid id)
    {
        await _lock.WaitAsync();
        try
        {
            await EnsureLoadedAsync();
            var pair = _config!.Pairs.FirstOrDefault(p => p.Id == id);
            if (pair != null)
            {
                _config.Pairs.Remove(pair);
                await SaveConfigAsync();
            }
        }
        finally
        {
            _lock.Release();
        }
    }

    private async Task EnsureLoadedAsync()
    {
        if (_config != null) return;

        await _lock.WaitAsync();
        try
        {
            if (_config != null) return;

            if (File.Exists(ConfigFile))
            {
                var json = await File.ReadAllTextAsync(ConfigFile);
                _config = JsonSerializer.Deserialize<TenantPairConfiguration>(json) ?? new TenantPairConfiguration();
            }
            else
            {
                _config = new TenantPairConfiguration();
            }
        }
        finally
        {
            _lock.Release();
        }
    }

    private async Task SaveConfigAsync()
    {
        Directory.CreateDirectory(ConfigFolder);
        var json = JsonSerializer.Serialize(_config, new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(ConfigFile, json);
    }
}

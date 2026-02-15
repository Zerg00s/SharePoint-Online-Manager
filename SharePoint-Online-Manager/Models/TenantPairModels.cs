namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents a source-to-target tenant pair for migration scenarios.
/// </summary>
public class TenantPair
{
    public Guid Id { get; set; } = Guid.NewGuid();

    /// <summary>
    /// The source connection ID (data migrated FROM).
    /// </summary>
    public Guid SourceConnectionId { get; set; }

    /// <summary>
    /// The target connection ID (data migrated TO).
    /// </summary>
    public Guid TargetConnectionId { get; set; }

    /// <summary>
    /// Optional display name for this pair.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// When this pair was created.
    /// </summary>
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// Configuration storage for tenant pairs.
/// </summary>
public class TenantPairConfiguration
{
    public List<TenantPair> Pairs { get; set; } = [];
}

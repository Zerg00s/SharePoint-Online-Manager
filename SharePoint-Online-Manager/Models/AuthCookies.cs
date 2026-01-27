namespace SharePointOnlineManager.Models;

/// <summary>
/// Represents the authentication cookies captured from SharePoint Online login.
/// </summary>
public class AuthCookies
{
    public string Domain { get; set; } = string.Empty;
    public string FedAuth { get; set; } = string.Empty;
    public string RtFa { get; set; } = string.Empty;
    public string UserEmail { get; set; } = string.Empty;
    public DateTime CapturedAt { get; set; } = DateTime.UtcNow;
    public DateTime? ExpiresAt { get; set; }

    public bool IsValid => !string.IsNullOrEmpty(FedAuth) && !string.IsNullOrEmpty(RtFa);

    /// <summary>
    /// Returns true if the cookies have expired.
    /// </summary>
    public bool IsExpired => ExpiresAt.HasValue && ExpiresAt.Value <= DateTime.UtcNow;

    /// <summary>
    /// Gets the time remaining until expiration, or null if no expiration is set.
    /// </summary>
    public TimeSpan? TimeRemaining => ExpiresAt.HasValue ? ExpiresAt.Value - DateTime.UtcNow : null;

    /// <summary>
    /// Gets a formatted string showing time remaining until expiration.
    /// </summary>
    public string TimeRemainingDisplay
    {
        get
        {
            if (!ExpiresAt.HasValue)
                return "Unknown";

            if (IsExpired)
                return "Expired";

            var remaining = TimeRemaining!.Value;
            if (remaining.TotalMinutes < 1)
                return "< 1 min";
            if (remaining.TotalHours < 1)
                return $"{(int)remaining.TotalMinutes} min";
            if (remaining.TotalHours < 24)
                return $"{(int)remaining.TotalHours}h {remaining.Minutes}m";

            return $"{(int)remaining.TotalDays}d {remaining.Hours}h";
        }
    }

    /// <summary>
    /// Gets the exact expiration date/time formatted for display.
    /// </summary>
    public string ExpirationDateTimeDisplay
    {
        get
        {
            if (!ExpiresAt.HasValue)
                return "Unknown";

            // Convert to local time for display
            var localExpiry = ExpiresAt.Value.ToLocalTime();
            return localExpiry.ToString("MMM dd HH:mm");
        }
    }

    /// <summary>
    /// Gets a combined display showing both exact time and remaining time.
    /// </summary>
    public string ExpirationDisplay
    {
        get
        {
            if (!ExpiresAt.HasValue)
                return "(re-auth to see)";

            if (IsExpired)
                return "EXPIRED";

            return $"{ExpirationDateTimeDisplay} ({TimeRemainingDisplay})";
        }
    }

    /// <summary>
    /// Gets the total duration of the token (from capture to expiration).
    /// </summary>
    public TimeSpan? TotalDuration => ExpiresAt.HasValue ? ExpiresAt.Value - CapturedAt : null;

    /// <summary>
    /// Gets a formatted string showing the total token duration in hours.
    /// </summary>
    public string TotalDurationDisplay
    {
        get
        {
            if (!TotalDuration.HasValue)
                return "-";

            var duration = TotalDuration.Value;
            if (duration.TotalHours < 1)
                return $"{(int)duration.TotalMinutes}m";
            if (duration.TotalHours < 48)
                return $"{duration.TotalHours:F0}h";

            return $"{duration.TotalDays:F1}d";
        }
    }
}

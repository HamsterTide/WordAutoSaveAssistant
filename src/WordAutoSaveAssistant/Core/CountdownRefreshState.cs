namespace WordAutoSaveAssistant.Core;

// Scheduling still ticks at 250 ms; only the visible countdown is coalesced.
internal sealed class CountdownRefreshState
{
    private int? _lastSeconds;

    internal bool TryGetChange(bool isVisible, TimeSpan? remaining, bool isSaving,
        out int totalSeconds, bool force = false)
    {
        totalSeconds = 0;
        if (!isVisible) return false;
        if (remaining is null || isSaving)
        {
            _lastSeconds = null;
            return false;
        }

        totalSeconds = Math.Max(0, (int)Math.Ceiling(remaining.Value.TotalSeconds));
        if (!force && totalSeconds == _lastSeconds) return false;
        _lastSeconds = totalSeconds;
        return true;
    }
}

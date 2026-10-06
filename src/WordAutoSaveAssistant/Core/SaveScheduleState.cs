namespace WordAutoSaveAssistant.Core;

public readonly record struct ScheduleAction(bool QueueSave, long Generation)
{
    public static ScheduleAction None(long generation) => new(false, generation);
}

public sealed class SaveScheduleState
{
    private readonly IMonotonicClock _clock;
    private long? _nextDue;
    private bool _pendingImmediate;

    public SaveScheduleState(IMonotonicClock clock, int intervalMinutes)
    {
        _clock = clock;
        SetIntervalValue(intervalMinutes);
    }

    public bool IsRunning { get; private set; }
    public bool IsSaving { get; private set; }
    public bool IsSessionLocked { get; private set; }
    public bool IsSystemSuspended { get; private set; }
    public bool IsPaused => IsSessionLocked || IsSystemSuspended;
    public int IntervalMinutes { get; private set; }
    public long Generation { get; private set; }

    public ScheduleAction Start(long now)
    {
        if (IsRunning)
        {
            return ScheduleAction.None(Generation);
        }

        IsRunning = true;
        Generation++;
        _nextDue = _clock.Add(now, Interval);

        if (IsPaused || IsSaving)
        {
            _pendingImmediate = true;
            return ScheduleAction.None(Generation);
        }

        IsSaving = true;
        return new ScheduleAction(true, Generation);
    }

    public void Stop()
    {
        if (!IsRunning && !IsSaving)
        {
            return;
        }

        IsRunning = false;
        Generation++;
        _nextDue = null;
        _pendingImmediate = false;
    }

    // A manual round never enables the timer or moves an existing deadline.
    public ScheduleAction SaveOnce()
    {
        if (IsSaving || IsPaused) return ScheduleAction.None(Generation);
        if (!IsRunning) Generation++;
        IsSaving = true;
        return new ScheduleAction(true, Generation);
    }

    public void ChangeInterval(int intervalMinutes, long now)
    {
        SetIntervalValue(intervalMinutes);
        if (IsRunning)
        {
            _nextDue = _clock.Add(now, Interval);
        }
    }

    public ScheduleAction Tick(long now)
    {
        if (!IsRunning || IsPaused || IsSaving || _nextDue is null || now < _nextDue.Value)
        {
            return ScheduleAction.None(Generation);
        }

        IsSaving = true;
        _nextDue = _clock.Add(now, Interval);
        return new ScheduleAction(true, Generation);
    }

    public ScheduleAction SetSessionLocked(bool locked, long now) =>
        SetPauseState(locked, IsSystemSuspended, now);

    public ScheduleAction SetSystemSuspended(bool suspended, long now) =>
        SetPauseState(IsSessionLocked, suspended, now);

    public ScheduleAction CompleteSave(long completedGeneration, long now)
    {
        IsSaving = false;

        if (!IsRunning || completedGeneration != Generation)
        {
            if (IsRunning && _pendingImmediate && !IsPaused)
            {
                _pendingImmediate = false;
                IsSaving = true;
                _nextDue = _clock.Add(now, Interval);
                return new ScheduleAction(true, Generation);
            }

            return ScheduleAction.None(Generation);
        }

        if (IsPaused)
        {
            _pendingImmediate = true;
            return ScheduleAction.None(Generation);
        }

        if (_pendingImmediate)
        {
            _pendingImmediate = false;
            IsSaving = true;
            _nextDue = _clock.Add(now, Interval);
            return new ScheduleAction(true, Generation);
        }

        if (_nextDue is null || now >= _nextDue.Value)
        {
            _nextDue = _clock.Add(now, Interval);
        }

        return ScheduleAction.None(Generation);
    }

    public TimeSpan? Remaining(long now)
    {
        if (!IsRunning || IsPaused || _nextDue is null)
        {
            return null;
        }

        TimeSpan remaining = _clock.Elapsed(now, _nextDue.Value);
        return remaining < TimeSpan.Zero ? TimeSpan.Zero : remaining;
    }

    private ScheduleAction SetPauseState(bool sessionLocked, bool systemSuspended, long now)
    {
        bool wasPaused = IsPaused;
        IsSessionLocked = sessionLocked;
        IsSystemSuspended = systemSuspended;
        bool isPaused = IsPaused;

        if (!IsRunning || wasPaused == isPaused)
        {
            return ScheduleAction.None(Generation);
        }

        if (isPaused)
        {
            _nextDue = null;
            _pendingImmediate = true;
            return ScheduleAction.None(Generation);
        }

        _nextDue = _clock.Add(now, Interval);
        if (IsSaving)
        {
            _pendingImmediate = true;
            return ScheduleAction.None(Generation);
        }

        _pendingImmediate = false;
        IsSaving = true;
        return new ScheduleAction(true, Generation);
    }

    private TimeSpan Interval => TimeSpan.FromMinutes(IntervalMinutes);

    private void SetIntervalValue(int intervalMinutes)
    {
        if (intervalMinutes is < 1 or > 1440)
        {
            throw new ArgumentOutOfRangeException(nameof(intervalMinutes), "间隔必须在 1 到 1440 分钟之间。");
        }

        IntervalMinutes = intervalMinutes;
    }
}

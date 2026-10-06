using System.Diagnostics;

namespace WordAutoSaveAssistant.Core;

public interface IMonotonicClock
{
    long Timestamp { get; }
    TimeSpan Elapsed(long start, long end);
    long Add(long timestamp, TimeSpan duration);
}

public sealed class MonotonicClock : IMonotonicClock
{
    public long Timestamp => Stopwatch.GetTimestamp();

    public TimeSpan Elapsed(long start, long end) =>
        TimeSpan.FromSeconds((end - start) / (double)Stopwatch.Frequency);

    public long Add(long timestamp, TimeSpan duration) =>
        timestamp + (long)(duration.TotalSeconds * Stopwatch.Frequency);
}

namespace WordAutoSaveAssistant.Services;

public sealed class SingleInstanceGuard : IDisposable
{
    private readonly Mutex _mutex;
    private bool _disposed;

    public SingleInstanceGuard(int sessionId)
    {
        string name = $"Local\\WordAutoSaveAssistant-{sessionId}";
        _mutex = new Mutex(true, name, out bool createdNew);
        IsPrimary = createdNew;
    }

    public bool IsPrimary { get; }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        if (IsPrimary)
        {
            try
            {
                _mutex.ReleaseMutex();
            }
            catch (ApplicationException)
            {
            }
        }

        _mutex.Dispose();
        _disposed = true;
    }
}

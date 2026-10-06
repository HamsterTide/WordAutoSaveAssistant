using System.Windows.Threading;

namespace WordAutoSaveAssistant.WordInterop;

internal sealed class ComStaHost
{
    private readonly Thread _thread;
    private readonly TaskCompletionSource<Dispatcher> _dispatcherReady =
        new(TaskCreationOptions.RunContinuationsAsynchronously);
    private Dispatcher? _dispatcher;
    private int _busy;
    private int _shutdownRequested;

    public ComStaHost()
    {
        _thread = new Thread(ThreadMain)
        {
            IsBackground = true,
            Name = "Word COM STA"
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public bool IsBusy => Volatile.Read(ref _busy) != 0;

    public async Task<T> InvokeAsync<T>(Func<CancellationToken, T> operation, CancellationToken cancellationToken)
    {
        if (Volatile.Read(ref _shutdownRequested) != 0)
        {
            throw new ObjectDisposedException(nameof(ComStaHost));
        }

        Dispatcher dispatcher = await _dispatcherReady.Task.ConfigureAwait(false);
        DispatcherOperation<T> dispatcherOperation = dispatcher.InvokeAsync(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            Interlocked.Exchange(ref _busy, 1);
            try
            {
                return operation(cancellationToken);
            }
            finally
            {
                Interlocked.Exchange(ref _busy, 0);
            }
        });

        return await dispatcherOperation.Task.ConfigureAwait(false);
    }

    public async Task<bool> TryShutdownAsync(TimeSpan timeout)
    {
        Interlocked.Exchange(ref _shutdownRequested, 1);
        Dispatcher dispatcher = await _dispatcherReady.Task.ConfigureAwait(false);

        try
        {
            dispatcher.BeginInvokeShutdown(DispatcherPriority.Send);
        }
        catch (InvalidOperationException)
        {
        }

        return await Task.Run(() => _thread.Join(timeout)).ConfigureAwait(false);
    }

    private void ThreadMain()
    {
        OleMessageFilter? messageFilter = null;
        try
        {
            _dispatcher = Dispatcher.CurrentDispatcher;
            messageFilter = new OleMessageFilter();
            messageFilter.Register();
            _dispatcherReady.TrySetResult(_dispatcher);
            Dispatcher.Run();
        }
        catch (Exception exception)
        {
            _dispatcherReady.TrySetException(exception);
        }
        finally
        {
            messageFilter?.Dispose();
        }
    }
}

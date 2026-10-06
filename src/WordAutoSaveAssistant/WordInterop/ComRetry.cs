using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WordAutoSaveAssistant.WordInterop;

internal static class ComRetry
{
    internal const int RpcCallRejected = unchecked((int)0x80010001);
    internal const int RpcServerCallRetryLater = unchecked((int)0x8001010A);

    public static T Execute<T>(Func<T> operation, CancellationToken token = default)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        while (true)
        {
            try
            {
                token.ThrowIfCancellationRequested();
                return operation();
            }
            catch (COMException exception) when (
                IsRetryable(exception.HResult) &&
                stopwatch.Elapsed < TimeSpan.FromSeconds(3))
            {
                if (token.WaitHandle.WaitOne(100)) token.ThrowIfCancellationRequested();
            }
        }
    }

    public static void Execute(Action operation, CancellationToken token = default) =>
        Execute(() =>
        {
            operation();
            return true;
        }, token);

    public static bool IsRetryable(int hResult) =>
        hResult is RpcCallRejected or RpcServerCallRetryLater;
}

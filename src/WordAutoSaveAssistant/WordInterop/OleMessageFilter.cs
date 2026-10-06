using System.Runtime.InteropServices;

namespace WordAutoSaveAssistant.WordInterop;

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);

    [PreserveSig]
    int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
}

internal sealed class OleMessageFilter : IOleMessageFilter, IDisposable
{
    private const int ServerCallRetryLater = 2;
    private const int RetryDelayMilliseconds = 100;
    private const int MaximumRetryMilliseconds = 3000;
    private bool _registered;

    public void Register()
    {
        int hr = CoRegisterMessageFilter(this, out _);
        Marshal.ThrowExceptionForHR(hr);
        _registered = true;
    }

    public int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo) => 0;

    public int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType) =>
        rejectType == ServerCallRetryLater && tickCount < MaximumRetryMilliseconds
            ? RetryDelayMilliseconds
            : -1;

    public int MessagePending(IntPtr taskCallee, int tickCount, int pendingType) => 2;

    public void Dispose()
    {
        if (!_registered)
        {
            return;
        }

        CoRegisterMessageFilter(null, out _);
        _registered = false;
    }

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        [MarshalAs(UnmanagedType.Interface)] IOleMessageFilter? newFilter,
        [MarshalAs(UnmanagedType.Interface)] out IOleMessageFilter? oldFilter);
}

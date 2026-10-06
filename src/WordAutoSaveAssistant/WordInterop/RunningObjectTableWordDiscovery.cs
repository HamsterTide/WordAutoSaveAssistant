using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using WordAutoSaveAssistant.Models;

namespace WordAutoSaveAssistant.WordInterop;

internal sealed record WordWindowCandidate(nint Handle, int ProcessId, string Title);

internal sealed record WordDiscoveryFailure(
    int ProcessId,
    string WindowTitle,
    string Message,
    string? ErrorCode,
    WordConnectionIssue ConnectionIssue = WordConnectionIssue.Unknown)
{
    public DocumentSaveOutcome ToOutcome() => new(
        string.IsNullOrWhiteSpace(WindowTitle) ? "(Word 实例)" : WindowTitle,
        DocumentSaveStatus.Failed,
        Message,
        ErrorCode,
        Source: DocumentOutcomeSource.WordInstance);
}

internal sealed record WordDiscoveryResult(
    IReadOnlyList<object> Applications,
    IReadOnlyList<WordDiscoveryFailure> Failures);

internal interface IWordWindowNativeApi
{
    IReadOnlyList<WordWindowCandidate> EnumerateWordWindows();
    object GetNativeObject(WordWindowCandidate candidate);
}

internal static class RunningObjectTableWordDiscovery
{
    public static WordDiscoveryResult DiscoverApplications(CancellationToken token = default) =>
        DiscoverApplications(new WordWindowNativeApi(), includeRotFallback: true, token: token);

    internal static WordDiscoveryResult RetryBusyProcess(int processId, CancellationToken token) =>
        DiscoverApplications(new WordWindowNativeApi(), false, processId, token);

    internal static WordDiscoveryResult DiscoverApplicationsForTest(IWordWindowNativeApi nativeApi) =>
        DiscoverApplications(nativeApi, includeRotFallback: false);

    private static WordDiscoveryResult DiscoverApplications(
        IWordWindowNativeApi nativeApi,
        bool includeRotFallback, int? processId = null, CancellationToken token = default)
    {
        ArgumentNullException.ThrowIfNull(nativeApi);
        Dictionary<long, object> applications = new();
        HashSet<int> accessibleProcessIds = new();
        List<WordDiscoveryFailure> failures = new();

        TryAddWindowApplications(nativeApi, applications, accessibleProcessIds, failures, processId, token);
        if (includeRotFallback && !token.IsCancellationRequested)
        {
            TryAddActiveApplication(applications, accessibleProcessIds);
            TryAddRotApplications(applications, accessibleProcessIds, token);
        }

        failures.RemoveAll(failure => accessibleProcessIds.Contains(failure.ProcessId));
        return new WordDiscoveryResult(applications.Values.ToArray(), failures);
    }

    private static void TryAddWindowApplications(
        IWordWindowNativeApi nativeApi,
        Dictionary<long, object> applications,
        HashSet<int> accessibleProcessIds,
        List<WordDiscoveryFailure> failures, int? processId, CancellationToken token)
    {
        IReadOnlyList<WordWindowCandidate> candidates;
        try
        {
            candidates = nativeApi.EnumerateWordWindows();
        }
        catch (Exception exception)
        {
            failures.Add(CreateFailure(
                new WordWindowCandidate(0, 0, "(Word 窗口枚举)"),
                exception));
            return;
        }

        foreach (WordWindowCandidate candidate in candidates)
        {
            if (token.IsCancellationRequested) break;
            if (processId.HasValue && candidate.ProcessId != processId.Value) continue;
            object? nativeObject = null;
            object? application = null;
            try
            {
                nativeObject = ComRetry.Execute(() => nativeApi.GetNativeObject(candidate));
                application = TryGetWordApplication(nativeObject)
                    ?? throw new InvalidOperationException("窗口没有返回可用的 Word 对象模型。");

                bool sameReference = ReferenceEquals(application, nativeObject);
                if (!TryStoreWordApplication(applications, application, out bool stored))
                {
                    throw new InvalidOperationException("窗口返回的对象不是可访问的 Word 应用程序。");
                }

                accessibleProcessIds.Add(candidate.ProcessId);
                if (!stored)
                {
                    ReleaseComObject(application);
                }
                if (sameReference)
                {
                    nativeObject = null;
                }
                application = null;
                if (!sameReference)
                {
                    ReleaseComObject(nativeObject);
                    nativeObject = null;
                }
            }
            catch (Exception exception)
            {
                if (application is not null && !ReferenceEquals(application, nativeObject))
                {
                    ReleaseComObject(application);
                }
                failures.Add(CreateFailure(candidate, exception));
            }
            finally
            {
                ReleaseComObject(nativeObject);
            }
        }
    }

    private static WordDiscoveryFailure CreateFailure(
        WordWindowCandidate candidate,
        Exception exception)
    {
        Exception effective = exception is AggregateException aggregateException
            ? aggregateException.GetBaseException()
            : exception;
        string title = string.IsNullOrWhiteSpace(candidate.Title)
            ? $"Word 窗口（PID {candidate.ProcessId}）"
            : candidate.Title;
        bool accessDenied = effective.HResult == unchecked((int)0x80070005);
        int? wordLevel = WordProcessSecurity.TryGetIntegrityLevel(candidate.ProcessId);
        int? assistantLevel = WordProcessSecurity.TryGetIntegrityLevel(Environment.ProcessId);
        WordConnectionIssue issue = WordProcessSecurity.ClassifyConnectionIssue(effective.HResult, wordLevel, assistantLevel);
        string boundary = WordProcessSecurity.DescribeConnectionBoundary(wordLevel, assistantLevel);
        string message = issue switch
        {
            WordConnectionIssue.PermissionMismatch => $"无法连接 Word 窗口“{title}”；权限差异已确认。{boundary}",
            WordConnectionIssue.Busy => $"无法连接 Word 窗口“{title}”；Word 暂时忙碌，请关闭阻塞对话框后重试",
            WordConnectionIssue.AccessDenied => $"无法连接 Word 窗口“{title}”；拒绝访问，权限级别可能不同，但差异尚未确认。{boundary}",
            WordConnectionIssue.ObjectModelUnavailable => $"无法连接 Word 窗口“{title}”；未公开桌面对象模型（或暂不可用）。{boundary}",
            _ => $"无法连接 Word 窗口“{title}”，原因未确认：{effective.Message}。{boundary}"
        };
        string? errorCode = effective is COMException || accessDenied
            ? $"0x{effective.HResult:X8}"
            : null;
        return new WordDiscoveryFailure(candidate.ProcessId, title, message, errorCode, issue);
    }

    private static void TryAddActiveApplication(
        Dictionary<long, object> applications,
        HashSet<int> accessibleProcessIds)
    {
        try
        {
            int classResult = CLSIDFromProgID("Word.Application", out Guid wordClassId);
            if (classResult < 0)
            {
                return;
            }

            int result = GetActiveObject(ref wordClassId, IntPtr.Zero, out object application);
            if (result >= 0 && TryStoreWordApplication(applications, application, out bool stored))
            {
                TryAddApplicationProcessId(application, accessibleProcessIds);
                if (!stored)
                {
                    ReleaseComObject(application);
                }
            }
            else if (result >= 0)
            {
                ReleaseComObject(application);
            }
        }
        catch
        {
            // GetActiveObject is only a fallback. Window failures are retained and shown.
        }
    }

    private static void TryAddRotApplications(
        Dictionary<long, object> applications,
        HashSet<int> accessibleProcessIds, CancellationToken token)
    {
        IRunningObjectTable? runningTable = null;
        IEnumMoniker? enumMoniker = null;
        try
        {
            Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out runningTable));
            runningTable.EnumRunning(out enumMoniker);
            enumMoniker.Reset();
            IMoniker[] monikers = new IMoniker[1];

            while (!token.IsCancellationRequested && enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
            {
                IMoniker? moniker = monikers[0];
                object? runningObject = null;
                object? application = null;
                try
                {
                    runningTable.GetObject(moniker, out runningObject);
                    application = TryGetWordApplication(runningObject);

                    if (application is null)
                    {
                        ReleaseComObject(runningObject);
                        runningObject = null;
                        continue;
                    }

                    bool sameReference = ReferenceEquals(application, runningObject);
                    if (TryStoreWordApplication(applications, application, out bool stored))
                    {
                        TryAddApplicationProcessId(application, accessibleProcessIds);
                        if (!stored)
                        {
                            ReleaseComObject(application);
                        }
                        if (sameReference)
                        {
                            runningObject = null;
                        }
                        application = null;
                    }
                    else
                    {
                        if (!sameReference)
                        {
                            ReleaseComObject(application);
                        }
                        application = null;
                        ReleaseComObject(runningObject);
                        runningObject = null;
                    }
                    if (!sameReference)
                    {
                        ReleaseComObject(runningObject);
                        runningObject = null;
                    }
                }
                catch
                {
                    if (application is not null && !ReferenceEquals(application, runningObject))
                    {
                        ReleaseComObject(application);
                    }
                    ReleaseComObject(runningObject);
                }
                finally
                {
                    ReleaseComObject(moniker);
                    monikers[0] = null!;
                }
            }
        }
        catch
        {
            // ROT is a best-effort fallback. Per-window failures remain visible.
        }
        finally
        {
            ReleaseComObject(enumMoniker);
            ReleaseComObject(runningTable);
        }
    }

    private static void TryAddApplicationProcessId(
        object application,
        HashSet<int> accessibleProcessIds)
    {
        try
        {
            dynamic candidate = application;
            nint windowHandle = (nint)Convert.ToInt64(
                ComRetry.Execute<object>(() => candidate.Hwnd));
            _ = GetWindowThreadProcessId(windowHandle, out uint processId);
            if (processId != 0)
            {
                accessibleProcessIds.Add(checked((int)processId));
            }
        }
        catch
        {
        }
    }

    private static object? TryGetWordApplication(object runningObject)
    {
        try
        {
            if (IsWordApplication(runningObject))
            {
                return runningObject;
            }
        }
        catch (COMException exception) when (ComRetry.IsRetryable(exception.HResult)) { throw; }
        catch
        {
        }

        object? application = null;
        try
        {
            dynamic candidate = runningObject;
            application = ComRetry.Execute<object>(() => candidate.Application);
            if (IsWordApplication(application))
            {
                return application;
            }
        }
        catch (COMException exception) when (ComRetry.IsRetryable(exception.HResult))
        {
            ReleaseComObject(application);
            throw;
        }
        catch
        {
        }

        ReleaseComObject(application);
        return null;
    }

    private static bool TryStoreWordApplication(
        Dictionary<long, object> applications,
        object application,
        out bool stored)
    {
        stored = false;
        try
        {
            if (!IsWordApplication(application))
            {
                return false;
            }

            long identity = GetComIdentity(application);
            if (applications.ContainsKey(identity))
            {
                return true;
            }

            applications.Add(identity, application);
            stored = true;
            return true;
        }
        catch (COMException exception) when (ComRetry.IsRetryable(exception.HResult)) { throw; }
        catch
        {
            return false;
        }
    }

    private static bool IsWordApplication(object application)
    {
        dynamic candidate = application;
        string name = Convert.ToString(ComRetry.Execute<object>(() => candidate.Name)) ?? string.Empty;
        // Do not probe Documents here: a busy collection must remain an
        // identifiable Word application so the save engine can defer it.
        return name.Contains("Word", StringComparison.OrdinalIgnoreCase);
    }

    public static long GetComIdentity(object comObject)
    {
        IntPtr unknown = Marshal.GetIUnknownForObject(comObject);
        try
        {
            return unknown.ToInt64();
        }
        finally
        {
            Marshal.Release(unknown);
        }
    }

    public static void ReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
        {
            return;
        }

        try
        {
            Marshal.ReleaseComObject(value);
        }
        catch
        {
        }
    }

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(
        int reserved,
        [MarshalAs(UnmanagedType.Interface)] out IRunningObjectTable runningObjectTable);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string progId, out Guid classId);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        ref Guid classId,
        IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object activeObject);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint windowHandle, out uint processId);
}

internal sealed class WordWindowNativeApi : IWordWindowNativeApi
{
    private const uint ObjIdNativeOm = 0xFFFFFFF0;
    private static readonly Guid IidDispatch = new("00020400-0000-0000-C000-000000000046");

    public IReadOnlyList<WordWindowCandidate> EnumerateWordWindows()
    {
        List<WordWindowCandidate> windows = new();
        EnumWindows((windowHandle, ignoredParameter) =>
        {
            if (!string.Equals(GetClassName(windowHandle), "OpusApp", StringComparison.Ordinal))
            {
                return true;
            }

            _ = GetWindowThreadProcessId(windowHandle, out uint processId);
            // Hidden owner/helper windows are not open document windows.
            string title = GetWindowTitle(windowHandle);
            if (!IsWindowVisible(windowHandle) || string.IsNullOrWhiteSpace(title)) return true;
            windows.Add(new WordWindowCandidate(
                windowHandle,
                checked((int)processId),
                title));
            return true;
        }, IntPtr.Zero);
        return windows;
    }

    public object GetNativeObject(WordWindowCandidate candidate)
    {
        List<nint> nativeObjectHandles = new();
        EnumChildWindows(candidate.Handle, (childHandle, ignoredParameter) =>
        {
            if (string.Equals(GetClassName(childHandle), "_WwG", StringComparison.Ordinal))
            {
                nativeObjectHandles.Add(childHandle);
            }
            return true;
        }, IntPtr.Zero);
        nativeObjectHandles.Add(candidate.Handle);

        Exception? lastError = null;
        foreach (nint windowHandle in nativeObjectHandles)
        {
            object? nativeObject = null;
            try
            {
                Guid dispatch = IidDispatch;
                int result = AccessibleObjectFromWindow(
                    windowHandle,
                    ObjIdNativeOm,
                    ref dispatch,
                    out nativeObject);
                Marshal.ThrowExceptionForHR(result);
                if (nativeObject is not null)
                {
                    return nativeObject;
                }
            }
            catch (Exception exception)
            {
                // Preserve the child object's access-denied error instead of
                // overwriting it with the top-level window's generic E_FAIL.
                if (lastError is null || exception.HResult == unchecked((int)0x80070005)) lastError = exception;
                RunningObjectTableWordDiscovery.ReleaseComObject(nativeObject);
            }
        }

        throw lastError ?? new InvalidOperationException("Word 窗口没有公开原生对象模型。");
    }

    private static string GetClassName(nint windowHandle)
    {
        StringBuilder className = new(128);
        _ = GetClassName(windowHandle, className, className.Capacity);
        return className.ToString();
    }

    private static string GetWindowTitle(nint windowHandle)
    {
        int length = GetWindowTextLength(windowHandle);
        if (length <= 0)
        {
            return string.Empty;
        }

        StringBuilder title = new(length + 1);
        _ = GetWindowText(windowHandle, title, title.Capacity);
        return title.ToString();
    }

    private delegate bool EnumWindowsProc(nint windowHandle, IntPtr parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(nint windowHandle);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumChildWindows(
        nint parentWindow,
        EnumWindowsProc callback,
        IntPtr parameter);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(
        nint windowHandle,
        StringBuilder className,
        int maximumCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(
        nint windowHandle,
        StringBuilder windowText,
        int maximumCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(nint windowHandle);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint windowHandle, out uint processId);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        nint windowHandle,
        uint objectId,
        ref Guid interfaceId,
        [MarshalAs(UnmanagedType.Interface)] out object nativeObject);
}

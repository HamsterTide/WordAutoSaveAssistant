using System.Runtime.InteropServices;
using System.Text;
using WordAutoSaveAssistant.WordInterop;

namespace WordAutoSaveAssistant.Tests;

internal static class WordDiscoveryDiagnostic
{
    public static int Run(IReadOnlyList<string> arguments)
    {
        foreach (System.Diagnostics.Process process in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
        {
            Console.WriteLine($"Word PID={process.Id} 完整性级别={WordProcessSecurity.TryGetIntegrityLevel(process.Id)?.ToString("X") ?? "无法读取"}；{WordProcessSecurity.DescribeConnectionBoundary(process.Id)}");
            process.Dispose();
        }
        if (arguments.Contains("--inventory", StringComparer.OrdinalIgnoreCase))
        {
            WordAutoSaveAssistant.Models.WordInventorySnapshot snapshot = new WordSaveEngine().InspectAll(CancellationToken.None);
            Console.WriteLine(snapshot.Summary);
            foreach (var item in snapshot.Items) Console.WriteLine($"  {item.DocumentName} | {item.State} | {item.FullPath} | {item.Detail}");
            return 0;
        }
        nint[] suppliedHandles = arguments
            .Where(argument => argument.StartsWith("--hwnd=", StringComparison.OrdinalIgnoreCase))
            .Select(argument => ParseHandle(argument[7..]))
            .ToArray();
        WordDiscoveryResult discovery = suppliedHandles.Length == 0
            ? RunningObjectTableWordDiscovery.DiscoverApplications()
            : RunningObjectTableWordDiscovery.DiscoverApplicationsForTest(
                new SuppliedWordWindowNativeApi(suppliedHandles));
        Console.WriteLine($"发现 Word 实例：{discovery.Applications.Count}");
        Console.WriteLine($"无法连接的 Word 窗口：{discovery.Failures.Count}");

        try
        {
            for (int applicationIndex = 0; applicationIndex < discovery.Applications.Count; applicationIndex++)
            {
                object application = discovery.Applications[applicationIndex];
                object? documents = null;
                try
                {
                    dynamic wordApplication = application;
                    documents = wordApplication.Documents;
                    dynamic collection = documents;
                    int count = Convert.ToInt32(collection.Count);
                    Console.WriteLine($"实例 {applicationIndex + 1}：{count} 个文档");
                    for (int documentIndex = 1; documentIndex <= count; documentIndex++)
                    {
                        object? document = null;
                        try
                        {
                            document = collection[documentIndex];
                            dynamic wordDocument = document;
                            Console.WriteLine($"  - {Convert.ToString(wordDocument.Name)}");
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine($"  - [无法读取名称] {Describe(exception)}");
                        }
                        finally
                        {
                            RunningObjectTableWordDiscovery.ReleaseComObject(document);
                        }
                    }
                }
                catch (Exception exception)
                {
                    Console.WriteLine($"实例 {applicationIndex + 1}：[无法读取] {Describe(exception)}");
                }
                finally
                {
                    RunningObjectTableWordDiscovery.ReleaseComObject(documents);
                }
            }

            foreach (WordDiscoveryFailure failure in discovery.Failures)
            {
                Console.WriteLine(
                    $"失败窗口 PID={failure.ProcessId} 标题={failure.WindowTitle} " +
                    $"错误={failure.ErrorCode ?? "无"} 原因={failure.Message}");
            }
        }
        finally
        {
            foreach (object application in discovery.Applications)
            {
                RunningObjectTableWordDiscovery.ReleaseComObject(application);
            }
        }

        return 0;
    }

    private static string Describe(Exception exception) =>
        exception is COMException comException
            ? $"0x{comException.HResult:X8} {comException.Message}"
            : exception.Message;

    private static nint ParseHandle(string value) =>
        value.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
            ? checked((nint)Convert.ToInt64(value[2..], 16))
            : checked((nint)Convert.ToInt64(value));

    private sealed class SuppliedWordWindowNativeApi : IWordWindowNativeApi
    {
        private readonly IReadOnlyList<WordWindowCandidate> _candidates;
        private readonly WordWindowNativeApi _nativeApi = new();

        public SuppliedWordWindowNativeApi(IEnumerable<nint> handles)
        {
            _candidates = handles.Select(handle =>
            {
                _ = GetWindowThreadProcessId(handle, out uint processId);
                return new WordWindowCandidate(
                    handle,
                    checked((int)processId),
                    ReadWindowTitle(handle));
            }).ToArray();
        }

        public IReadOnlyList<WordWindowCandidate> EnumerateWordWindows() => _candidates;

        public object GetNativeObject(WordWindowCandidate candidate) =>
            _nativeApi.GetNativeObject(candidate);

        private static string ReadWindowTitle(nint handle)
        {
            int length = GetWindowTextLength(handle);
            StringBuilder title = new(Math.Max(1, length + 1));
            _ = GetWindowText(handle, title, title.Capacity);
            return title.Length == 0 ? $"HWND 0x{handle:X}" : title.ToString();
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(nint windowHandle, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(
            nint windowHandle,
            StringBuilder windowText,
            int maximumCount);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(nint windowHandle);
    }
}

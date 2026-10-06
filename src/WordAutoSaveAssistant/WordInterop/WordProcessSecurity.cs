using System.Runtime.InteropServices;
using WordAutoSaveAssistant.Models;

namespace WordAutoSaveAssistant.WordInterop;

internal static class WordProcessSecurity
{
    public static int? TryGetIntegrityLevel(int processId)
    {
        nint process = OpenProcess(0x1000, false, processId); // Query limited information only.
        if (process == 0) return null;
        nint token = 0;
        nint buffer = 0;
        try
        {
            if (!OpenProcessToken(process, 0x0008, out token)) return null;
            GetTokenInformation(token, 25, 0, 0, out int size); // TokenIntegrityLevel.
            if (size <= 0) return null;
            buffer = Marshal.AllocHGlobal(size);
            if (!GetTokenInformation(token, 25, buffer, size, out _)) return null;
            nint sid = Marshal.ReadIntPtr(buffer);
            int count = Marshal.ReadByte(GetSidSubAuthorityCount(sid));
            return count == 0 ? null : Marshal.ReadInt32(GetSidSubAuthority(sid, (uint)(count - 1)));
        }
        catch { return null; }
        finally
        {
            if (buffer != 0) Marshal.FreeHGlobal(buffer);
            if (token != 0) CloseHandle(token);
            CloseHandle(process);
        }
    }

    public static string DescribeConnectionBoundary(int processId)
    {
        int? word = TryGetIntegrityLevel(processId);
        int? assistant = TryGetIntegrityLevel(Environment.ProcessId);
        return DescribeConnectionBoundary(word, assistant);
    }

    internal static WordConnectionIssue ClassifyConnectionIssue(int hresult, int? word, int? assistant)
    {
        if (ComRetry.IsRetryable(hresult)) return WordConnectionIssue.Busy;
        if (word is not null && assistant is not null && word != assistant)
            return WordConnectionIssue.PermissionMismatch;
        return hresult switch
        {
            unchecked((int)0x80070005) => WordConnectionIssue.AccessDenied,
            unchecked((int)0x80004005) => WordConnectionIssue.ObjectModelUnavailable,
            _ => WordConnectionIssue.Unknown
        };
    }

    internal static string DescribeConnectionBoundary(int? word, int? assistant)
    {
        if (word is not null && assistant is not null && word != assistant)
            return $"Word 与助手的权限级别不同（Word：{LevelName(word.Value)}；助手：{LevelName(assistant.Value)}）；请手动以相同权限运行两者";
        if (word is null || assistant is null)
            return "权限信息未能完整读取，连接失败原因尚未确认";
        return "两者权限级别一致，不应直接归因于权限；请检查阻塞对话框或 Word 的运行状态";
    }

    private static string LevelName(int level) => level >= 0x3000 ? "高权限" : level >= 0x2000 ? "普通权限" : "受限权限";

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern nint OpenProcess(uint access, bool inherit, int processId);
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool OpenProcessToken(nint process, uint access, out nint token);
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool GetTokenInformation(nint token, int informationClass, nint information, int length, out int required);
    [DllImport("advapi32.dll")]
    private static extern nint GetSidSubAuthorityCount(nint sid);
    [DllImport("advapi32.dll")]
    private static extern nint GetSidSubAuthority(nint sid, uint index);
    [DllImport("kernel32.dll")]
    private static extern bool CloseHandle(nint handle);
}

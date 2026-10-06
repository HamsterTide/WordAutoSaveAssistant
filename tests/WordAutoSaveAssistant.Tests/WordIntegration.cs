using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO.Compression;
using System.Windows.Threading;
using WordAutoSaveAssistant.Core;
using WordAutoSaveAssistant.Models;
using WordAutoSaveAssistant.WordInterop;

namespace WordAutoSaveAssistant.Tests;

internal static class WordIntegration
{
    public static int Run(bool timed = false)
    {
        HashSet<int> existingWordProcesses = Process.GetProcessesByName("WINWORD")
            .Select(process => process.Id)
            .ToHashSet();
        string root = Path.Combine(Path.GetTempPath(), $"word-save-assistant-integration-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        object? application = null;
        List<object> documents = new();
        int? ownedProcessId = null;

        try
        {
            Type? wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType is null)
            {
                Console.WriteLine("SKIP: 未安装可自动化的 Microsoft Word。");
                return 2;
            }

            application = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("无法创建隔离 Word 实例。");
            dynamic word = application;
            HashSet<int> newWordProcesses = new();
            for (int attempt = 0; attempt < 50; attempt++)
            {
                newWordProcesses = Process.GetProcessesByName("WINWORD")
                    .Select(process => process.Id)
                    .Where(processId => !existingWordProcesses.Contains(processId))
                    .ToHashSet();
                if (newWordProcesses.Count > 0)
                {
                    break;
                }
                Thread.Sleep(100);
            }

            if (newWordProcesses.Count != 1)
            {
                Console.WriteLine("SKIP: Word 自动化连接到了启动测试前已经存在的进程；为保护用户文档，未执行任何文档操作。");
                return 2;
            }
            ownedProcessId = newWordProcesses.Single();

            word.Visible = false;
            word.DisplayAlerts = 0;

            dynamic docs = word.Documents;
            string firstPath = Path.Combine(root, "changed-a.docx");
            string secondPath = Path.Combine(root, "changed-b.docx");
            string unchangedPath = Path.Combine(root, "unchanged.docx");
            string readOnlyPath = Path.Combine(root, "read-only.docx");
            string protectedPath = Path.Combine(root, "protected.docx");

            dynamic first = docs.Add();
            documents.Add(first);
            first.Content.Text = "初始 A";
            first.SaveAs2(firstPath);
            first.Content.InsertAfter("—修改");

            dynamic second = docs.Add();
            documents.Add(second);
            second.Content.Text = "初始 B";
            second.SaveAs2(secondPath);
            second.Content.InsertAfter("—修改");

            dynamic unchanged = docs.Add();
            documents.Add(unchanged);
            unchanged.Content.Text = "无需再次保存";
            unchanged.SaveAs2(unchangedPath);

            dynamic unsaved = docs.Add();
            documents.Add(unsaved);
            unsaved.Content.Text = "没有保存路径";

            dynamic readOnlySeed = docs.Add();
            readOnlySeed.Content.Text = "只读测试";
            readOnlySeed.SaveAs2(readOnlyPath);
            readOnlySeed.Close(0);
            ReleaseComObject(readOnlySeed);
            dynamic readOnly = docs.Open(readOnlyPath, ReadOnly: true);
            documents.Add(readOnly);

            dynamic protectedDocument = docs.Add();
            documents.Add(protectedDocument);
            protectedDocument.Content.Text = "受保护测试";
            protectedDocument.SaveAs2(protectedPath);
            protectedDocument.Protect(2);

            SaveRoundResult result = new WordSaveEngine().ProcessKnownApplicationsForTest(
                new[] { application },
                CancellationToken.None);

            bool passed = result.DetectedInstances == 1
                && result.DetectedDocuments == 6
                && result.SavedCount == 2
                && result.UnchangedCount == 1
                && result.SkippedCount == 3
                && result.FailedCount == 0
                && result.Outcomes.All(outcome => outcome.Source == DocumentOutcomeSource.Document)
                && result.Outcomes.Where(outcome => outcome.Status == DocumentSaveStatus.Saved)
                    .All(outcome => !string.IsNullOrWhiteSpace(outcome.FullPath))
                && result.Outcomes.Where(outcome => outcome.Status == DocumentSaveStatus.Unchanged)
                    .All(outcome => !string.IsNullOrWhiteSpace(outcome.FullPath))
                && result.Outcomes.Count(outcome => outcome.Status == DocumentSaveStatus.Skipped
                    && string.IsNullOrWhiteSpace(outcome.FullPath)) == 1
                && result.Outcomes.Count(outcome => outcome.Status == DocumentSaveStatus.Skipped
                    && !string.IsNullOrWhiteSpace(outcome.FullPath)) == 2;

            Console.WriteLine(
                $"Word 隔离测试：实例 {result.DetectedInstances}，文档 {result.DetectedDocuments}，" +
                $"已保存 {result.SavedCount}，无修改 {result.UnchangedCount}，跳过 {result.SkippedCount}，失败 {result.FailedCount}");
            foreach (DocumentSaveOutcome outcome in result.Outcomes)
            {
                Console.WriteLine($"  {outcome.Status}: {outcome.DocumentName} — {outcome.Message} — {outcome.FullPath ?? "(无路径)"}");
            }

            passed &= DiskContains(firstPath, "—修改") && DiskContains(secondPath, "—修改");
            // Exercise manual mode against real Word without enabling the timer.
            ((dynamic)first).Content.InsertAfter("—单次保存验证");
            ((dynamic)second).Content.InsertAfter("—单次保存验证");
            MonotonicClock manualClock = new();
            SaveScheduleState manualSchedule = new(manualClock, 10);
            ScheduleAction manualAction = manualSchedule.SaveOnce();
            SaveRoundResult manualResult = new WordSaveEngine().ProcessKnownApplicationsForTest([application], CancellationToken.None);
            manualSchedule.CompleteSave(manualAction.Generation, manualClock.Timestamp);
            bool manualPassed = manualAction.QueueSave && !manualSchedule.IsRunning && !manualSchedule.IsSaving
                && manualSchedule.Remaining(manualClock.Timestamp) is null && manualResult.SavedCount == 2
                && DiskContains(firstPath, "—单次保存验证") && DiskContains(secondPath, "—单次保存验证");
            passed &= manualPassed;
            Console.WriteLine(manualPassed ? "PASS: 单次保存落盘后定时仍停止" : "FAIL: 单次保存验证");
            if (passed && timed)
            {
                passed &= RunTimedCheck(application, first, second, firstPath, secondPath);
            }
            Console.WriteLine(passed ? "PASS: Word 隔离保存及磁盘内容验证" : "FAIL: Word 隔离保存或磁盘内容验证");
            return passed ? 0 : 1;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"FAIL: Word 隔离测试异常：{exception}");
            return 1;
        }
        finally
        {
            foreach (object document in documents.AsEnumerable().Reverse())
            {
                try { ((dynamic)document).Close(0); } catch { }
                ReleaseComObject(document);
            }

            if (application is not null && ownedProcessId is int processId && !existingWordProcesses.Contains(processId))
            {
                try { ((dynamic)application).Quit(0); } catch { }
            }
            ReleaseComObject(application);
            try { Directory.Delete(root, true); } catch { }
        }
    }

    private static bool RunTimedCheck(object application, object first, object second, string firstPath, string secondPath)
    {
        MonotonicClock clock = new();
        SaveScheduleState schedule = new(clock, 1);
        schedule.Start(clock.Timestamp);
        schedule.CompleteSave(schedule.Generation, clock.Timestamp);
        ((dynamic)first).Content.InsertAfter("—一分钟定时验证");
        ((dynamic)second).Content.InsertAfter("—一分钟定时验证");
        if (Convert.ToBoolean(((dynamic)first).Saved) || Convert.ToBoolean(((dynamic)second).Saved))
            throw new InvalidOperationException("测试文档未进入有修改状态。");

        Dispatcher dispatcher = Dispatcher.CurrentDispatcher;
        ComStaHost host = new();
        SaveRoundResult? timedResult = null;
        Exception? failure = null;
        Stopwatch elapsed = Stopwatch.StartNew();
        DispatcherTimer timer = new(DispatcherPriority.Background) { Interval = TimeSpan.FromMilliseconds(250) };
        timer.Tick += async (_, _) =>
        {
            if (elapsed.Elapsed > TimeSpan.FromSeconds(90))
            {
                failure = new TimeoutException("一分钟定时验证超时。");
                timer.Stop();
                dispatcher.BeginInvokeShutdown(DispatcherPriority.Background);
                return;
            }
            ScheduleAction action = schedule.Tick(clock.Timestamp);
            if (!action.QueueSave) return;
            timer.Stop();
            try
            {
                timedResult = await host.InvokeAsync(token => new WordSaveEngine()
                    .ProcessKnownApplicationsForTest(new[] { application }, token), CancellationToken.None);
                schedule.CompleteSave(action.Generation, clock.Timestamp);
                schedule.Stop();
            }
            catch (Exception exception) { failure = exception; }
            finally { dispatcher.BeginInvokeShutdown(DispatcherPriority.Background); }
        };
        Console.WriteLine("一分钟定时验证：两份临时文档已再次编辑，等待真实 DispatcherTimer 到期；仅操作测试专用 Word。");
        timer.Start();
        Dispatcher.Run();
        host.TryShutdownAsync(TimeSpan.FromSeconds(3)).GetAwaiter().GetResult();
        if (failure is not null) throw failure;
        bool passed = elapsed.Elapsed >= TimeSpan.FromSeconds(60)
            && timedResult is { SavedCount: 2, FailedCount: 0 }
            && DiskContains(firstPath, "一分钟定时验证") && DiskContains(secondPath, "一分钟定时验证");
        Console.WriteLine($"{(passed ? "PASS" : "FAIL")}: 实际 {elapsed.Elapsed.TotalSeconds:F1} 秒后触发，已保存 {timedResult?.SavedCount}，两份磁盘 DOCX 均包含新内容。");
        return passed;
    }

    private static bool DiskContains(string path, string expected)
    {
        using FileStream file = new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using ZipArchive package = new(file, ZipArchiveMode.Read);
        using StreamReader reader = new(package.GetEntry("word/document.xml")!.Open());
        return reader.ReadToEnd().Contains(expected, StringComparison.Ordinal);
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.FinalReleaseComObject(value); } catch { }
        }
    }

}

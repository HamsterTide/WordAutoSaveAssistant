using System.Runtime.InteropServices;
using WordAutoSaveAssistant.Models;

namespace WordAutoSaveAssistant.WordInterop;

internal sealed class WordSaveEngine
{
    private readonly Action<TimeSpan, CancellationToken> _delay;
    private readonly Func<int, CancellationToken, WordDiscoveryResult> _retryProcess;
    internal WordSaveEngine(Action<TimeSpan, CancellationToken>? delay = null,
        Func<int, CancellationToken, WordDiscoveryResult>? retryProcess = null)
    {
        _delay = delay ?? ((duration, token) =>
        {
            if (token.WaitHandle.WaitOne(duration)) token.ThrowIfCancellationRequested();
        });
        _retryProcess = retryProcess ?? RunningObjectTableWordDiscovery.RetryBusyProcess;
    }

    public SaveRoundResult SaveAll(CancellationToken token) => SaveAndInspectAll(token).Round;

    public WordScanResult SaveAndInspectAll(CancellationToken token, IProgress<string>? progress = null) =>
        ProcessApplications(RunningObjectTableWordDiscovery.DiscoverApplications(token), true, true, token, progress);

    public WordInventorySnapshot InspectAll(CancellationToken token) =>
        ProcessApplications(RunningObjectTableWordDiscovery.DiscoverApplications(token), true, false, token).Inventory;

    internal SaveRoundResult ProcessKnownApplicationsForTest(IReadOnlyList<object> applications, CancellationToken token) =>
        ProcessApplications(new(applications, []), false, true, token).Round;

    internal SaveRoundResult ProcessDiscoveryForTest(WordDiscoveryResult discovery, CancellationToken token) =>
        ProcessApplications(discovery, false, true, token).Round;

    internal WordInventorySnapshot InspectKnownApplicationsForTest(IReadOnlyList<object> applications, CancellationToken token) =>
        ProcessApplications(new(applications, []), false, false, token).Inventory;

    internal WordInventorySnapshot InspectDiscoveryForTest(WordDiscoveryResult discovery, CancellationToken token) =>
        ProcessApplications(discovery, false, false, token).Inventory;

    private WordScanResult ProcessApplications(
        WordDiscoveryResult discovery, bool releaseApplications, bool saveChanges, CancellationToken token,
        IProgress<string>? progress = null)
    {
        List<DocumentSaveOutcome> outcomes = new();
        List<WordInventoryItem> inventory = new();
        List<object> heldDocuments = new();
        List<object> retryApplications = new();
        List<Action> delayed = new();
        using ComIdentityScope identities = new();
        using ComIdentityScope applicationIdentities = new();
        int instanceCount = 0;
        int documentCount = 0;
        bool cancelled = false;

        try
        {
            foreach (WordDiscoveryFailure failure in discovery.Failures)
                if (!saveChanges || failure.ProcessId <= 0 || !IsBusy(failure.ErrorCode)) AddWindowFailure(failure);
            if (saveChanges)
                foreach (var group in discovery.Failures.Where(f => f.ProcessId > 0 && IsBusy(f.ErrorCode)).GroupBy(f => f.ProcessId))
                    delayed.Add(() =>
                    {
                        WordDiscoveryResult retry = _retryProcess(group.Key, token);
                        retryApplications.AddRange(retry.Applications);
                        if (retry.Applications.Count == 0 && retry.Failures.Count == 0)
                            foreach (WordDiscoveryFailure original in group) AddWindowFailure(original);
                        foreach (WordDiscoveryFailure failure in retry.Failures) AddWindowFailure(failure);
                        foreach (object app in retry.Applications) Capture(app, false);
                    });

            // Capture all initial references before any Save changes collections.
            foreach (object application in discovery.Applications)
            {
                token.ThrowIfCancellationRequested();
                Capture(application, saveChanges);
            }

            int initialCount = heldDocuments.Count;
            foreach (object document in heldDocuments)
            {
                token.ThrowIfCancellationRequested();
                Evaluate(document, saveChanges);
            }
            if (delayed.Count > 0)
            {
                progress?.Report("Word 正忙，5 秒后重试一次；可点击停止取消等待。");
                _delay(TimeSpan.FromSeconds(5), token);
                token.ThrowIfCancellationRequested();
                progress?.Report("正在重试忙碌项（仅一次）…");
                foreach (Action action in delayed)
                {
                    token.ThrowIfCancellationRequested();
                    action();
                }
                for (int index = initialCount; index < heldDocuments.Count; index++)
                {
                    token.ThrowIfCancellationRequested();
                    Evaluate(heldDocuments[index], false, retried: true);
                }
            }
        }
        catch (OperationCanceledException) { cancelled = true; }
        finally
        {
            foreach (object document in heldDocuments) RunningObjectTableWordDiscovery.ReleaseComObject(document);
            foreach (object app in retryApplications) RunningObjectTableWordDiscovery.ReleaseComObject(app);
            if (releaseApplications)
                foreach (object app in discovery.Applications) RunningObjectTableWordDiscovery.ReleaseComObject(app);
        }

        DateTimeOffset now = DateTimeOffset.Now;
        return new(
            new(now, instanceCount, documentCount, outcomes, cancelled),
            new(now, instanceCount, documentCount,
                inventory.OrderBy(item => item.Source)
                    .ThenBy(item => item.DocumentName, StringComparer.CurrentCultureIgnoreCase)
                    .ThenBy(item => item.FullPath, StringComparer.OrdinalIgnoreCase).ToArray(), cancelled));

        void Capture(object application, bool allowDelay)
        {
            if (applicationIdentities.TryAdd(application)) instanceCount++;
            object? documents = null;
            bool deferred = false;
            try
            {
                dynamic app = application;
                documents = ComRetry.Execute<object>(() => app.Documents, token);
                dynamic collection = documents;
                int count = Convert.ToInt32(ComRetry.Execute<object>(() => collection.Count, token));
                for (int index = 1; index <= count; index++)
                {
                    token.ThrowIfCancellationRequested();
                    object? document = null;
                    try
                    {
                        document = ComRetry.Execute<object>(() => collection[index], token);
                        if (identities.TryAdd(document))
                        {
                            heldDocuments.Add(document);
                            document = null;
                            documentCount++;
                        }
                    }
                    catch (Exception exception) when (exception is not OperationCanceledException)
                    {
                        if (allowDelay && ComRetry.IsRetryable(exception.HResult)) deferred = true;
                        else AddFailure($"(第 {index} 个文档无法读取)", exception, DocumentOutcomeSource.Document);
                    }
                    finally { RunningObjectTableWordDiscovery.ReleaseComObject(document); }
                }
            }
            catch (Exception exception) when (exception is not OperationCanceledException)
            {
                if (allowDelay && ComRetry.IsRetryable(exception.HResult)) deferred = true;
                else AddFailure("(Word 实例)", exception, DocumentOutcomeSource.WordInstance);
            }
            finally { RunningObjectTableWordDiscovery.ReleaseComObject(documents); }
            if (deferred) delayed.Add(() => Capture(application, false));
        }

        void Evaluate(object document, bool allowDelay, bool retried = false, bool saveReturned = false)
        {
            DocumentEvaluation evaluation = ProcessDocument(document, saveChanges, token, saveReturned);
            if (allowDelay && IsBusy(evaluation.Outcome?.ErrorCode))
            {
                // If Save already returned, retry confirmation only, not the write.
                delayed.Add(() => Evaluate(document, false, true, evaluation.SaveReturned));
                return;
            }
            if (retried)
            {
                const string suffix = "（Word 忙碌后已延迟重试一次）";
                evaluation = new(evaluation.Inventory with { Detail = evaluation.Inventory.Detail + suffix },
                    evaluation.Outcome is null ? null : evaluation.Outcome with { Message = evaluation.Outcome.Message + suffix });
            }
            inventory.Add(evaluation.Inventory);
            if (evaluation.Outcome is not null) outcomes.Add(evaluation.Outcome);
        }

        void AddWindowFailure(WordDiscoveryFailure failure)
        {
            outcomes.Add(failure.ToOutcome());
            inventory.Add(new(failure.WindowTitle, null, "无法连接", failure.Message + $" {failure.ErrorCode}",
                false, null, DocumentOutcomeSource.WordInstance, ConnectionIssue: failure.ConnectionIssue));
        }

        void AddFailure(string name, Exception exception, DocumentOutcomeSource source)
        {
            DocumentSaveOutcome failure = Failure(name, exception, source);
            outcomes.Add(failure);
            inventory.Add(new(name, null, "无法读取", failure.Message + $" {failure.ErrorCode}", false, null, source,
                ConnectionIssue: ComRetry.IsRetryable(exception.HResult) ? WordConnectionIssue.Busy : WordConnectionIssue.Unknown));
        }
    }

    private static bool IsBusy(string? code) => code is "0x80010001" or "0x8001010A";

    private sealed record DocumentEvaluation(WordInventoryItem Inventory, DocumentSaveOutcome? Outcome, bool SaveReturned = false);

    private static DocumentEvaluation ProcessDocument(object document, bool saveChanges, CancellationToken token, bool saveReturned = false)
    {
        dynamic doc = document;
        string name;
        try { name = Convert.ToString(ComRetry.Execute<object>(() => doc.Name, token)) ?? "(未命名文档)"; }
        catch (OperationCanceledException) { throw; }
        catch { name = "(无法读取名称)"; }
        string? fullPath = null;
        try
        {
            string path = Convert.ToString(ComRetry.Execute<object>(() => doc.Path, token)) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(path))
            {
                try { fullPath = Convert.ToString(ComRetry.Execute<object>(() => doc.FullName, token)); }
                catch (OperationCanceledException) { throw; }
                catch { /* A missing tooltip path must not prevent a normal Save. */ }
            }
            if (string.IsNullOrWhiteSpace(path))
                return Skip("未命名文档没有现有保存路径", "未命名 · 不自动保存");
            if (Convert.ToBoolean(ComRetry.Execute<object>(() => doc.ReadOnly, token)))
                return Skip("文档为只读状态", "只读 · 不自动保存");
            if (Convert.ToInt32(ComRetry.Execute<object>(() => doc.ProtectionType, token)) != -1)
                return Skip("文档处于受保护状态", "受保护 · 不自动保存");

            bool saved = Convert.ToBoolean(ComRetry.Execute<object>(() => doc.Saved, token));
            if (saved)
                return new(new(name, fullPath, "无未保存修改", "Word 报告没有未保存的修改", true, false),
                    saveChanges ? new(name, saveReturned ? DocumentSaveStatus.Saved : DocumentSaveStatus.Unchanged,
                        saveReturned ? "保存成功，延迟读取后已确认" : "没有未保存的修改", FullPath: fullPath) : null);
            if (!saveChanges)
                return new(new(name, fullPath, "待保存", "检测到未保存的修改；可立即保存一次或开始定时保存", true, true), null);

            // Retry each call separately rather than replaying the entire Save flow.
            if (!saveReturned)
            {
                ComRetry.Execute(() => { doc.Save(); }, token);
                saveReturned = true;
            }
            saved = Convert.ToBoolean(ComRetry.Execute<object>(() => doc.Saved, token));
            if (saved)
                return new(new(name, fullPath, "无未保存修改", "本轮已保存，Word 已确认", true, false),
                    new(name, DocumentSaveStatus.Saved, "保存成功", FullPath: fullPath));
            const string pending = "已调用保存，但 Word 仍报告未保存修改（可能仍在编辑或后台保存）";
            return new(new(name, fullPath, "未确认保存", pending, true, true),
                new(name, DocumentSaveStatus.Failed, pending, FullPath: fullPath));
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            DocumentSaveOutcome failure = Failure(name, exception, DocumentOutcomeSource.Document, fullPath);
            return new(new(name, fullPath, "无法读取 / 保存", failure.Message + $" {failure.ErrorCode}", false, null), failure, saveReturned);
        }

        DocumentEvaluation Skip(string reason, string state) => new(
            new(name, fullPath, state, reason, false, null, IsSkipped: true),
            saveChanges ? new(name, DocumentSaveStatus.Skipped, reason, FullPath: fullPath) : null);
    }

    private static DocumentSaveOutcome Failure(string name, Exception exception, DocumentOutcomeSource source, string? fullPath = null)
    {
        string message = exception switch
        {
            COMException when ComRetry.IsRetryable(exception.HResult) => "Word 正忙，本轮未能完成；请关闭对话框后重试",
            COMException => $"Word 返回错误：{exception.Message}",
            _ => $"处理文档时发生错误：{exception.Message}"
        };
        return new(name, DocumentSaveStatus.Failed, message, $"0x{exception.HResult:X8}", fullPath, source);
    }
}
